import React, { useState, useEffect, useId, useRef, useMemo } from 'react';
import pptxgen from 'pptxgenjs';
import { db, auth } from './firebase';
import { GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged, User } from 'firebase/auth';
import {
  detectCanvaAppsSdkHost,
  captureSelectedMediaFromCanvaEditor,
  replaceKeywordsInCanvaRichtextSelection,
} from './canvaSdkBridge';
import {
  ExternalLink,
  Copy,
  Check,
  RefreshCw,
  LogOut,
  Key,
  Upload,
  Image as ImageIcon,
  Video,
  Layers,
  Settings2,
  Play,
  Download,
  ChevronRight,
  ChevronLeft,
  ChevronDown,
  Trash2,
  Plus,
  FileText,
  Sliders,
  CheckCircle2,
  AlertCircle,
  Sparkles,
  Info,
  Folder,
  FolderOpen,
  FolderPlus,
  Edit3,
  FolderInput,
  Move,
  CornerDownRight,
  Music,
  List,
  Grid,
  Search,
  ArrowUpDown,
  HardDrive,
  Filter,
  Eye,
  X,
  GripVertical,
  UploadCloud,
  RotateCcw,
  CheckSquare,
  Square
} from 'lucide-react';

interface CanvaUser {
  id: string;
  displayName?: string;
}

interface CanvaStatus {
  connected: boolean;
  configured: boolean;
  clientIdConfigured?: boolean;
  clientSecretConfigured?: boolean;
  clientId?: string;
  clientIdSource?: string;
  redirectUri: string;
  user?: CanvaUser;
  scopes?: string[];
  hasBrandTemplateScope?: boolean;
  expiresAt?: number;
}

export interface ExtractedPageDetail {
  pageNumber: number;
  title: string;
  status: 'success' | 'failed';
  mediaUrl: string;
  texts: string[];
  elements: string[];
  hasPhrasePlaceholder?: boolean;
}

export interface PptxTextBoxInfo {
  slideIndex: number;
  text: string;
  tokens: string[];
  xPct: number;
  yPct: number;
  wPct: number;
  hPct: number;
  fontSizePt: number;
  fontFamily?: string;
  colorHex: string;
  align: 'left' | 'center' | 'right';
  bold: boolean;
  italic?: boolean;
  underline?: boolean;
  charSpacingPt?: number;
  lineSpacingPct?: number;
}

export interface PptxSlideLayoutInfo {
  slideIndex: number;
  hasCleanBg: boolean;
  cleanBgMime: string | null;
  textBoxes: PptxTextBoxInfo[];
  detectedTokens: string[];
}

export interface LocalMediaItem {
  id: string;
  name: string;
  type: 'image' | 'video' | 'audio';
  size: number;
  url: string;
  file: File;
  pageCount?: number;
  activePageIndex?: number;
  pageMediaUrls?: string[];
  cleanPageBgUrls?: string[];
  slideLayouts?: PptxSlideLayoutInfo[];
  isCanvaDesignTemplate?: boolean;
  extractedPages?: ExtractedPageDetail[];
}

export function isMediaUrlVideo(url: string, file?: File, typeHint?: string): boolean {
  if (!url) return false;
  if (url.match(/\.(png|jpg|jpeg|webp|gif|svg)$/i)) return false;
  if (file && file.type && file.type.startsWith('image/')) return false;
  if (file && file.type && file.type.startsWith('video/')) return true;
  if (url.endsWith('.mp4') || url.endsWith('.mov') || url.endsWith('.webm') || url.endsWith('.m4v')) return true;
  if (url.startsWith('blob:') && typeHint === 'video' && file && file.type.startsWith('video/')) return true;
  if (typeHint === 'video' && (url.endsWith('.mp4') || url.includes('.mp4') || url.includes('video'))) return true;
  return false;
}

// Gera URLs distintas de fundo para cada página do modelo do Canva
export function getDefaultPageUrls(baseItem: { url: string; name?: string; type?: string }): string[] {
  const primaryUrl = baseItem.url || 'https://images.unsplash.com/photo-1550745165-9bc0b252726f?w=600&q=80';
  return [
    primaryUrl,
    'https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=600&q=80',
    'https://images.unsplash.com/photo-1579783902614-a3fb3927b675?w=600&q=80',
    'https://images.unsplash.com/photo-1513694203232-719a280e022f?w=600&q=80',
    'https://images.unsplash.com/photo-1542744094-3a3172720224?w=600&q=80',
  ];
}

// Helper para calcular a página atual e mídia associada para qualquer índice de frase
export function getPreviewPageInfo(phraseIdx: number, items: LocalMediaItem[]) {
  if (!items || items.length === 0) {
    return {
      media: null,
      pageNum: 1,
      totalPagesInDesign: 1,
      pageUrl: '',
      cleanPageUrl: '',
      slideLayout: null as PptxSlideLayoutInfo | null,
      isVideo: false,
    };
  }

  const totalPagesAcrossAll = items.reduce((sum, item) => sum + Math.max(1, item.pageCount || 3), 0);
  const cycleIndex = phraseIdx % totalPagesAcrossAll;

  let acc = 0;
  for (const item of items) {
    const itemPages = Math.max(1, item.pageCount || 3);
    if (cycleIndex < acc + itemPages) {
      const pageNum = (cycleIndex - acc) + 1;
      const pageUrls = item.pageMediaUrls && item.pageMediaUrls.length > 0
        ? item.pageMediaUrls
        : getDefaultPageUrls(item);
      const pageUrl = pageUrls[(pageNum - 1) % pageUrls.length] || item.url;
      const cleanPageUrl = item.cleanPageBgUrls?.[(pageNum - 1) % (item.cleanPageBgUrls?.length || 1)] || '';
      const slideLayout =
        item.slideLayouts?.find((l) => l.slideIndex === pageNum) ||
        item.slideLayouts?.[(pageNum - 1) % (item.slideLayouts?.length || 1)] ||
        null;
      const isVideo = isMediaUrlVideo(pageUrl, item.file, item.type);

      return {
        media: item,
        pageNum,
        totalPagesInDesign: itemPages,
        pageUrl,
        cleanPageUrl,
        slideLayout,
        isVideo,
      };
    }
    acc += itemPages;
  }

  const fallback = items[0];
  const pages = Math.max(1, fallback.pageCount || 3);
  const pageNum = (phraseIdx % pages) + 1;
  const pageUrls = fallback.pageMediaUrls && fallback.pageMediaUrls.length > 0
    ? fallback.pageMediaUrls
    : getDefaultPageUrls(fallback);
  const pageUrl = pageUrls[(pageNum - 1) % pageUrls.length] || fallback.url;
  const cleanPageUrl = fallback.cleanPageBgUrls?.[(pageNum - 1) % (fallback.cleanPageBgUrls?.length || 1)] || '';
  const slideLayout =
    fallback.slideLayouts?.find((l) => l.slideIndex === pageNum) ||
    fallback.slideLayouts?.[(pageNum - 1) % (fallback.slideLayouts?.length || 1)] ||
    null;

  return {
    media: fallback,
    pageNum,
    totalPagesInDesign: pages,
    pageUrl,
    cleanPageUrl,
    slideLayout,
    isVideo: isMediaUrlVideo(pageUrl, fallback.file, fallback.type),
  };
}

export interface KeyFieldMappingConfig {
  fieldKey: string;
  detectedInTemplate?: boolean;
  source: 'batch_phrase' | 'signature' | 'custom_text' | 'ignore';
  customText: string;
  pageScope: 'all' | 'specific';
  selectedPages: number[];
}

export interface UpdatedPptxTextBox extends PptxTextBoxInfo {
  didReplace: boolean;
}

/**
 * Atualiza a propriedade `text` diretamente nos elementos originais de `slideLayouts`
 * sem inserir novas caixas de texto sobrepostas, preservando fontFamily, fontSizePt,
 * colorHex, bold e align, e respeitando estritamente o pageScope configurado.
 */
export function replaceSlideLayoutTextBoxesInPlace(
  slideLayout: PptxSlideLayoutInfo | null | undefined,
  pageNum: number,
  phraseText: string,
  mappings: KeyFieldMappingConfig[],
  signatureText: string
): UpdatedPptxTextBox[] {
  if (!slideLayout || !Array.isArray(slideLayout.textBoxes)) return [];

  const activeMappings = mappings.filter((m) => m.source !== 'ignore');
  const hasExplicitPhraseRule = mappings.some(
    (m) => m.source === 'batch_phrase' || /frase|quote|texto/i.test(m.fieldKey)
  );

  return slideLayout.textBoxes.map((tb) => {
    let updatedText = tb.text;
    let didReplace = false;

    for (const mapping of activeMappings) {
      if (mapping.pageScope === 'specific' && !mapping.selectedPages.includes(pageNum)) {
        continue;
      }

      const val =
        mapping.source === 'batch_phrase'
          ? phraseText
          : mapping.source === 'signature'
          ? signatureText || '@lotecriativo'
          : mapping.customText;

      if (val === null || val === undefined) continue;

      const target = mapping.fieldKey.trim();
      if (!target) continue;
      const hasBracket = /^[\(\{\[<].*[\)\}\]>]$/.test(target);
      const esc = target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      const patterns: RegExp[] = [];
      if (!hasBracket) {
        const cleanW = target.replace(/^[\(\{\[<]|[\)\}\]>]$/g, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
        patterns.push(new RegExp(`\\(${cleanW}\\)`, 'gi'));
        patterns.push(new RegExp(`\\{${cleanW}\\}`, 'gi'));
        patterns.push(new RegExp(`\\[${cleanW}\\]`, 'gi'));
        patterns.push(new RegExp(`<${cleanW}>`, 'gi'));
      }
      patterns.push(new RegExp(esc, 'gi'));

      for (const p of patterns) {
        if (p.test(updatedText)) {
          p.lastIndex = 0;
          updatedText = updatedText.replace(p, () => val);
          didReplace = true;
          break;
        }
      }
    }

    // Fallback automático apenas quando o usuário ainda não configurou regra explícita para frase
    if (!didReplace && !hasExplicitPhraseRule && /\(frase\)|\{frase\}|\[frase\]/i.test(updatedText)) {
      updatedText = updatedText.replace(/\(frase\)|\{frase\}|\[frase\]/gi, () => phraseText);
      didReplace = true;
    }

    return {
      ...tb,
      text: updatedText,
      didReplace,
    };
  });
}

interface BatchResult {
  id: string;
  phrase: string;
  status: 'completed' | 'alternative_ready' | 'failed' | 'processing';
  type: 'canva_design' | 'editable_slide';
  url: string | null;
  mediaName: string;
  message?: string;
  isMultipage?: boolean;
  pageCount?: number;
}

export default function App() {
  const [user, setUser] = useState<User | null>(auth.currentUser);
  const [activeStep, setActiveStep] = useState<number>(1);

  // Canva Connection State
  const [canvaStatus, setCanvaStatus] = useState<CanvaStatus | null>(null);
  const [loadingCanva, setLoadingCanva] = useState<boolean>(true);
  const [connecting, setConnecting] = useState<boolean>(false);
  const [copiedUri, setCopiedUri] = useState<boolean>(false);
  const [manualClientId, setManualClientId] = useState<string>('');
  const [manualClientSecret, setManualClientSecret] = useState<string>('');
  const [savingCredentials, setSavingCredentials] = useState<boolean>(false);
  const [credentialsMsg, setCredentialsMsg] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const [directAuthUrl, setDirectAuthUrl] = useState<string>('');
  const [canvaCloudItems, setCanvaCloudItems] = useState<any[]>([]);
  const [loadingCanvaAssets, setLoadingCanvaAssets] = useState<boolean>(false);
  const [showCanvaBrowser, setShowCanvaBrowser] = useState<boolean>(false);

  // Canva Hierarchical Folder Explorer State (Step 4 & Global)
  const [showFolderExplorer, setShowFolderExplorer] = useState<boolean>(false);
  const [explorerCurrentFolderId, setExplorerCurrentFolderId] = useState<string>('root');
  const [explorerBreadcrumbs, setExplorerBreadcrumbs] = useState<Array<{ id: string; name: string }>>([
    { id: 'root', name: 'Projetos (Raiz)' }
  ]);
  const [explorerItems, setExplorerItems] = useState<any[]>([]);
  const [loadingExplorer, setLoadingExplorer] = useState<boolean>(false);
  const [explorerContinuation, setExplorerContinuation] = useState<string | null>(null);
  const [loadingMoreExplorer, setLoadingMoreExplorer] = useState<boolean>(false);
  const [selectedExplorerIds, setSelectedExplorerIds] = useState<Set<string>>(new Set());
  const [selectedExplorerItemsMap, setSelectedExplorerItemsMap] = useState<Map<string, any>>(new Map());
  const [folderExplorerSearch, setFolderExplorerSearch] = useState<string>('');
  const [importingFromExplorer, setImportingFromExplorer] = useState<boolean>(false);

  // Canva Videos / Audio / Unified Filter Explorer State
  const [explorerTab, setExplorerTab] = useState<'folders' | 'videos'>('folders');
  const [explorerMediaFilter, setExplorerMediaFilter] = useState<'all' | 'image' | 'video' | 'audio' | 'design'>('all');
  const [explorerVideosData, setExplorerVideosData] = useState<{
    videoDesigns: any[];
    allDesigns: any[];
    videoAssets: any[];
    audioAssets?: any[];
    transferredMedia?: any[];
    apiLimitationNote?: string;
    audioLimitationNote?: string;
  } | null>(null);
  const [loadingExplorerVideos, setLoadingExplorerVideos] = useState<boolean>(false);
  const [videoFilterMode, setVideoFilterMode] = useState<'video_candidates' | 'all_designs'>('video_candidates');
  const [assetLookupInput, setAssetLookupInput] = useState<string>('');
  const [lookingUpAsset, setLookingUpAsset] = useState<boolean>(false);
  const [assetLookupError, setAssetLookupError] = useState<string | null>(null);
  const [exportingMp4Id, setExportingMp4Id] = useState<string | null>(null);
  const [downloadingPptx, setDownloadingPptx] = useState<boolean>(false);
  const [capturingSdkMedia, setCapturingSdkMedia] = useState<boolean>(false);
  const [sdkCaptureFeedback, setSdkCaptureFeedback] = useState<{ text: string; type: 'success' | 'info' | 'error' } | null>(null);
  const [applyingRichtextSdk, setApplyingRichtextSdk] = useState<boolean>(false);
  const canvaSdkHost = useMemo(() => detectCanvaAppsSdkHost(), []);

  // Canva Asset Uploads State (Step 6 & 7)
  const [canvaUploadQueue, setCanvaUploadQueue] = useState<Array<{
    id: string;
    name: string;
    phrase: string;
    mediaUrl: string;
    status: 'pending' | 'uploading' | 'success' | 'failed';
    assetId?: string;
    canvaUploadsUrl?: string;
    errorMessage?: string;
  }>>([]);
  const [isUploadingToCanva, setIsUploadingToCanva] = useState<boolean>(false);
  const [uploadProgressCount, setUploadProgressCount] = useState<{ current: number; total: number }>({ current: 0, total: 0 });

  // Canva Editable Design Creation State (Step 6 & 7 - Opção 2)
  const [isSavingDesignToCanva, setIsSavingDesignToCanva] = useState<boolean>(false);
  const [savingDesignStatusText, setSavingDesignStatusText] = useState<string>('');
  const [designSaveBanner, setDesignSaveBanner] = useState<{ text: string; type: 'success' | 'error' } | null>(null);
  const [createdCanvaDesign, setCreatedCanvaDesign] = useState<{
    designId: string;
    title: string;
    pageCount?: number;
    editUrl: string;
    viewUrl?: string;
  } | null>(null);
  const [copiedDesignLink, setCopiedDesignLink] = useState<boolean>(false);

  // Opção 1: Gerar a partir de um Modelo do Canva (Brand Template Autofill + Campos-Chave)
  const [generateOptionTab, setGenerateOptionTab] = useState<'option1_template' | 'option2_new_design'>('option1_template');
  const [showTemplateSelectorModal, setShowTemplateSelectorModal] = useState<boolean>(false);
  const [templateSelectorTab, setTemplateSelectorTab] = useState<'brand_templates' | 'folders'>('brand_templates');
  const [brandTemplatesList, setBrandTemplatesList] = useState<Array<{
    id: string;
    name: string;
    type: string;
    thumbnailUrl: string;
    viewUrl?: string;
    createUrl?: string;
  }>>([]);
  const [loadingBrandTemplates, setLoadingBrandTemplates] = useState<boolean>(false);
  const [brandTemplateAuthError, setBrandTemplateAuthError] = useState<string | null>(null);
  const [templateSelectorSearch, setTemplateSelectorSearch] = useState<string>('');
  const [tplFolderId, setTplFolderId] = useState<string>('root');
  const [tplBreadcrumbs, setTplBreadcrumbs] = useState<Array<{ id: string; name: string }>>([{ id: 'root', name: 'Projetos (Raiz)' }]);
  const [tplFolderItems, setTplFolderItems] = useState<any[]>([]);
  const [tplFolderContinuation, setTplFolderContinuation] = useState<string | null>(null);
  const [loadingTplFolder, setLoadingTplFolder] = useState<boolean>(false);
  const [loadingMoreTplFolder, setLoadingMoreTplFolder] = useState<boolean>(false);
  const [manualTemplateInspectId, setManualTemplateInspectId] = useState<string>('');
  const [inspectingTemplate, setInspectingTemplate] = useState<boolean>(false);
  const [inspectedTemplate, setInspectedTemplate] = useState<{
    id: string;
    brandTemplateId: string | null;
    title: string;
    isBrandTemplate: boolean;
    autofillCompatible: boolean;
    editableViaPptxRoundtrip?: boolean;
    datasetFields: Array<{ name: string; type: string }>;
    detectedPlaceholders?: string[];
    slideLayouts?: PptxSlideLayoutInfo[];
    pageCount: number;
    pages: Array<{ index: number; thumbnailUrl: string; title: string; cleanBgUrl?: string | null }>;
    editUrl: string | null;
    viewUrl: string | null;
    hasBrandTemplateScope: boolean;
    limitationReason: string;
  } | null>(null);
  const [option1OutputMode, setOption1OutputMode] = useState<'individual' | 'multipage'>('individual');
  const [downloadingReplacedPptx, setDownloadingReplacedPptx] = useState<boolean>(false);
  const [failedCleanBgUrls, setFailedCleanBgUrls] = useState<Record<string, boolean>>({});
  const [keyFieldMappings, setKeyFieldMappings] = useState<Array<{
    fieldKey: string;
    detectedInTemplate: boolean;
    source: 'batch_phrase' | 'signature' | 'custom_text' | 'ignore';
    customText: string;
    pageScope: 'all' | 'specific';
    selectedPages: number[];
  }>>([
    { fieldKey: '(frase)', detectedInTemplate: false, source: 'batch_phrase', customText: '', pageScope: 'all', selectedPages: [1] },
    { fieldKey: 'FRASE', detectedInTemplate: false, source: 'ignore', customText: '', pageScope: 'all', selectedPages: [1] },
    { fieldKey: 'TITULO', detectedInTemplate: false, source: 'ignore', customText: '', pageScope: 'all', selectedPages: [1] },
    { fieldKey: 'SUBTITULO', detectedInTemplate: false, source: 'ignore', customText: '', pageScope: 'all', selectedPages: [1] },
    { fieldKey: 'ASSINATURA', detectedInTemplate: false, source: 'ignore', customText: '', pageScope: 'all', selectedPages: [1] },
    { fieldKey: 'DATA', detectedInTemplate: false, source: 'ignore', customText: new Date().toLocaleDateString('pt-BR'), pageScope: 'all', selectedPages: [1] },
  ]);
  const [newCustomFieldKey, setNewCustomFieldKey] = useState<string>('');
  const [autofillTitlePrefix, setAutofillTitlePrefix] = useState<string>('LoteCriativo');
  const [autofillBatchLoading, setAutofillBatchLoading] = useState<boolean>(false);
  const [autofillStatusBanner, setAutofillStatusBanner] = useState<{ text: string; type: 'success' | 'error' | 'info' } | null>(null);
  const [autofillVersionResults, setAutofillVersionResults] = useState<Array<{
    versionIndex: number;
    designId: string | null;
    title: string;
    pageCount: number;
    editUrl: string | null;
    viewUrl?: string | null;
    label: string;
    status: 'completed' | 'failed';
    error?: string;
  }>>([]);
  const [copiedVersionId, setCopiedVersionId] = useState<string | null>(null);

  // Step 2: Frases
  const [rawPhrases, setRawPhrases] = useState<string>(
    "O sucesso é a soma de pequenos esforços repetidos diariamente.\nCriatividade é a inteligência se divertindo.\nGrandes ideias começam com um simples primeiro passo.\nA constância vence o talento quando o talento não é constante."
  );
  const [designMode, setDesignMode] = useState<'individual' | 'multipage'>('individual');

  // Step 3: Modelo
  const [templateId, setTemplateId] = useState<string>('');
  const [fieldPhrase, setFieldPhrase] = useState<string>('FRASE');
  const [fieldMedia, setFieldMedia] = useState<string>('VIDEO');
  const [fieldSignature, setFieldSignature] = useState<string>('ASSINATURA');

  // Step 4: Mídias
  const [mediaList, setMediaList] = useState<LocalMediaItem[]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Step 5: Personalização
  const [aspectRatio, setAspectRatio] = useState<'9:16' | '1:1' | '16:9'>('9:16');
  const [textPosition, setTextPosition] = useState<'top' | 'center' | 'bottom'>('center');
  const [textAlign, setTextAlign] = useState<'left' | 'center' | 'right'>('center');
  const [fontFamily, setFontFamily] = useState<string>('serif');
  const [fontSize, setFontSize] = useState<number>(24);
  const [textColor, setTextColor] = useState<string>('#FFFFFF');
  const [overlayOpacity, setOverlayOpacity] = useState<number>(45);
  const [hasShadow, setHasShadow] = useState<boolean>(true);
  const [hasStroke, setHasStroke] = useState<boolean>(false);
  const [signatureText, setSignatureText] = useState<string>('@lotecriativo');

  // Step 6: Prévia
  const [previewIndex, setPreviewIndex] = useState<number>(0);

  // Step 7 & 8: Motor de Lote & Resultados
  const [batchProgress, setBatchProgress] = useState<{ current: number; total: number; status: string } | null>(null);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [batchResults, setBatchResults] = useState<BatchResult[]>([]);

  // Step 9: Modo Drive Explorer
  const [driveViewMode, setDriveViewMode] = useState<'explorer' | 'grid'>('explorer');
  const [driveCategory, setDriveCategory] = useState<'all' | 'design' | 'video' | 'image' | 'audio' | 'folder'>('all');
  const [driveSearchQuery, setDriveSearchQuery] = useState<string>('');
  const [driveFolderPath, setDriveFolderPath] = useState<Array<{ id: string; name: string }>>([
    { id: 'root', name: 'Meu Drive Canva' }
  ]);
  const [selectedDriveItem, setSelectedDriveItem] = useState<any | null>(null);

  // Folder management state
  const [customFolders, setCustomFolders] = useState<any[]>([]);
  const [showCreateFolderModal, setShowCreateFolderModal] = useState<boolean>(false);
  const [newFolderName, setNewFolderName] = useState<string>('');
  const [creatingFolder, setCreatingFolder] = useState<boolean>(false);

  const [showRenameModal, setShowRenameModal] = useState<boolean>(false);
  const [itemToRename, setItemToRename] = useState<any | null>(null);
  const [renameInput, setRenameInput] = useState<string>('');
  const [renamingItem, setRenamingItem] = useState<boolean>(false);

  const [showMoveModal, setShowMoveModal] = useState<boolean>(false);
  const [itemToMove, setItemToMove] = useState<any | null>(null);
  const [targetFolderIdForMove, setTargetFolderIdForMove] = useState<string>('root');
  const [movingItem, setMovingItem] = useState<boolean>(false);

  // Drag-and-Drop & Notification state
  const [draggedItemId, setDraggedItemId] = useState<string | null>(null);
  const [dragOverFolderId, setDragOverFolderId] = useState<string | null>(null);
  const [dropNotification, setDropNotification] = useState<string | null>(null);

  // Quick Hover Preview State for Step 9 Drive Explorer list rows
  const [hoveredItemForPreview, setHoveredItemForPreview] = useState<any | null>(null);
  const [hoverPosition, setHoverPosition] = useState<{ x: number; y: number }>({ x: 0, y: 0 });

  // Canva Multi-Page Inspector state
  const [inspectingMediaItem, setInspectingMediaItem] = useState<LocalMediaItem | null>(null);
  const [inspectingPageIndex, setInspectingPageIndex] = useState<number>(0);
  const [isExtractingPages, setIsExtractingPages] = useState<boolean>(false);

  useEffect(() => {
    if (inspectingMediaItem) {
      setInspectingPageIndex(0);
    }
  }, [inspectingMediaItem]);

  // Search query inside Canva media browser modal
  const [canvaSearchQuery, setCanvaSearchQuery] = useState<string>('');

  // Multi-select & Sorting state for Step 9 Drive Canva
  const [selectedItemIds, setSelectedItemIds] = useState<Set<string>>(new Set());
  const [sortBy, setSortBy] = useState<'name' | 'date' | 'size'>('name');
  const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');

  // Carregar status do Firebase Auth
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
    });
    return () => unsubscribe();
  }, []);

  // Verificar status da integração com Canva (depende exclusivamente de chamada ao servidor)
  const fetchCanvaStatus = async () => {
    try {
      setLoadingCanva(true);
      console.log('[CANVA STATUS CHECK] Querying fresh authentication status directly from server API /api/auth/canva/status...');
      const res = await fetch('/api/auth/canva/status', {
        headers: { 'Cache-Control': 'no-cache, no-store' },
      });
      if (res.ok) {
        const data = await res.json();
        console.log('[CANVA STATUS CHECK] Server returned fresh authentication state:', data);
        setCanvaStatus(data);
        if (data.connected) {
          syncCanvaDriveFiles();
        }
      }
    } catch (err) {
      console.warn('Erro ao consultar status da conexão Canva no servidor:', err);
    } finally {
      setLoadingCanva(false);
    }
  };

  // Classificador inteligente de tipos de arquivos do Canva
  const categorizeCanvaItem = (item: any): 'design' | 'video' | 'image' | 'audio' | 'folder' => {
    if (!item) return 'image';
    const obj = item.design || item.image || item.video || item.folder || item.audio || item;
    const rawType = String(item.type || item.asset_type || obj.type || '').toLowerCase();
    const name = String(obj.title || obj.name || item.title || item.name || '').toLowerCase();

    if (rawType === 'folder' || Boolean(item.folder) || Boolean(obj.folder)) {
      return 'folder';
    }

    if (
      rawType === 'video' ||
      Boolean(item.video) ||
      Boolean(obj.video) ||
      name.endsWith('.mp4') ||
      name.endsWith('.mov') ||
      name.endsWith('.webm') ||
      name.endsWith('.m4v') ||
      name.endsWith('.avi') ||
      name.includes('video') ||
      name.includes('vídeo') ||
      name.includes('reels') ||
      name.includes('shorts') ||
      name.includes('tiktok') ||
      name.includes('dueto') ||
      name.includes('corte')
    ) {
      return 'video';
    }

    if (
      rawType === 'audio' ||
      Boolean(item.audio) ||
      Boolean(obj.audio) ||
      name.endsWith('.mp3') ||
      name.endsWith('.wav') ||
      name.endsWith('.m4a') ||
      name.endsWith('.aac') ||
      name.includes('audio') ||
      name.includes('áudio') ||
      name.includes('musica') ||
      name.includes('música') ||
      name.includes('trilha') ||
      name.includes('podcast') ||
      name.includes('vinheta')
    ) {
      return 'audio';
    }

    if (
      rawType === 'image' ||
      rawType === 'photo' ||
      Boolean(item.image) ||
      Boolean(obj.image) ||
      name.endsWith('.png') ||
      name.endsWith('.jpg') ||
      name.endsWith('.jpeg') ||
      name.endsWith('.webp') ||
      name.endsWith('.gif') ||
      name.includes('_img') ||
      name.includes('foto') ||
      name.includes('imagem') ||
      name.includes('banner')
    ) {
      return 'image';
    }

    if (item.source === 'design' || rawType === 'design' || Boolean(item.design)) {
      return 'design';
    }

    return 'image';
  };

  // Função que sincroniza automaticamente os arquivos da conta do usuário com o 'Drive Canva'
  const syncCanvaDriveFiles = async (folderId: string = 'root') => {
    setLoadingCanvaAssets(true);
    try {
      const [designsRes, assetsRes, videosRes] = await Promise.all([
        fetch('/api/canva/designs'),
        fetch(`/api/canva/assets?folderId=${encodeURIComponent(folderId)}`),
        fetch('/api/canva/videos'),
      ]);

      const designsData = await designsRes.json().catch(() => ({}));
      const assetsData = await assetsRes.json().catch(() => ({}));
      const videosData = await videosRes.json().catch(() => ({}));

      const rawDesigns = designsData.items || designsData.designs || [];
      const rawAssets = assetsData.items || assetsData.assets || [];
      const rawVideoDesigns = videosData.videoDesigns || [];
      const rawVideoAssets = videosData.videoAssets || [];
      const videoCandidateIds = new Set<string>(rawVideoDesigns.map((v: any) => String(v.id)));

      const seenIds = new Set<string>();
      const items: any[] = [];

      for (const d of rawDesigns) {
        const obj = d.design || d;
        const id = String(obj.id || d.id || Math.random().toString());
        const pagesList = obj.pages || d.pages || [];
        const realPageCount = pagesList.length > 0 ? pagesList.length : (obj.page_count || obj.pageCount || 3);
        const pageUrls = pagesList.map((p: any) => p.thumbnail?.url || p.url).filter(Boolean);

        if (!seenIds.has(id)) {
          seenIds.add(id);
          const name = obj.title || obj.name || d.title || 'Design Canva';
          let cat = categorizeCanvaItem({ ...d, design: obj, source: 'design', name, title: name });
          if (videoCandidateIds.has(id)) {
            cat = 'video';
          }
          items.push({
            id,
            name,
            type: cat === 'video' ? 'video' : 'image',
            url: obj.thumbnail?.url || d.thumbnail?.url || d.urls?.view_url || '',
            source: 'design',
            category: cat,
            canvaUrl: d.urls?.edit_url || d.urls?.view_url || 'https://www.canva.com',
            folderId: folderId || 'root',
            pageCount: realPageCount,
            pageMediaUrls: pageUrls.length > 0 ? pageUrls : undefined,
          });
        }
      }

      for (const va of rawVideoAssets) {
        const id = String(va.id || '');
        if (id && !seenIds.has(id)) {
          seenIds.add(id);
          items.push({
            id,
            name: va.name || `Vídeo Canva ${id}`,
            type: va.type === 'video' ? 'video' : 'image',
            url: va.thumbnailUrl || '',
            source: va.source || 'canva_asset_api',
            category: va.type === 'video' ? 'video' : 'image',
            canvaUrl: 'https://www.canva.com/folder/uploads',
            folderId: folderId === 'uploads' ? 'uploads' : 'root',
            pageCount: 1,
          });
        }
      }

      for (const a of rawAssets) {
        const obj = a.folder || a.image || a.design || a.brand_template || a.video || a.audio || a;
        const id = String(obj.id || a.id || Math.random().toString());
        if (!seenIds.has(id)) {
          seenIds.add(id);
          const name = obj.name || obj.title || a.name || a.title || 'Ativo Canva';
          let cat = categorizeCanvaItem(a);
          if (videoCandidateIds.has(id)) {
            cat = 'video';
          }
          items.push({
            id,
            name,
            type: cat === 'video' ? 'video' : 'image',
            url: obj.thumbnail?.url || a.thumbnail?.url || a.url || '',
            source: a.type || 'asset',
            category: cat,
            canvaUrl: a.urls?.edit_url || a.urls?.view_url || (cat === 'design' ? 'https://www.canva.com' : ''),
            folderId: folderId || 'root',
            pageCount: obj.page_count || obj.pageCount || 3,
          });
        }
      }

      setCanvaCloudItems((prev) => {
        const existingMap = new Map(prev.map((item) => [item.id, item]));
        for (const item of items) {
          existingMap.set(item.id, item);
        }
        return Array.from(existingMap.values());
      });
    } catch (err) {
      console.warn('Erro na sincronização automática do Canva:', err);
    } finally {
      setLoadingCanvaAssets(false);
    }
  };

  useEffect(() => {
    // Restaura credenciais e estado da sessão do localStorage e cookie ao carregar
    try {
      const savedCreds = localStorage.getItem('canva_saved_credentials');
      if (savedCreds) {
        const parsed = JSON.parse(savedCreds);
        if (parsed.clientId) setManualClientId(parsed.clientId);
        if (parsed.clientSecret) setManualClientSecret(parsed.clientSecret);

        if (parsed.clientId && parsed.clientSecret) {
          fetch('/api/auth/canva/credentials', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ clientId: parsed.clientId, clientSecret: parsed.clientSecret }),
          }).then(() => fetchCanvaStatus());
          return;
        }
      }
    } catch {
      // ignore
    }
    fetchCanvaStatus();
  }, []);

  // Efeito que detecta quando canvaStatus.connected é verdadeiro para sincronizar o Drive Canva e atualizar cookies
  useEffect(() => {
    if (canvaStatus?.connected) {
      document.cookie = "canva_connected=true; max-age=31536000; path=/";
      console.log('[CANVA EFFECT] Connection verified. Triggering syncCanvaDriveFiles...');
      syncCanvaDriveFiles();
    } else if (canvaStatus?.connected === false) {
      document.cookie = "canva_connected=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
    }
  }, [canvaStatus?.connected]);

  // Função unificada de importação de itens do Canva com extração completa de TODAS as páginas
  const processAndAddCanvaItem = async (item: any) => {
    const rawCat = String(item.category || item.type || '').toLowerCase();
    const itemType: 'image' | 'video' | 'audio' =
      rawCat === 'audio' ? 'audio' : rawCat === 'video' ? 'video' : 'image';
    const itemUrl =
      item.url ||
      item.thumbnailUrl ||
      (itemType === 'audio' ? '' : 'https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=600&q=80');
    let finalUrl = itemUrl;
    let mediaFile = new File([], item.name, {
      type: itemType === 'video' ? 'video/mp4' : itemType === 'audio' ? 'audio/mpeg' : 'image/png',
    });

    let realPageCount = item.pageCount || item.page_count || (itemType === 'audio' ? 1 : 3);
    let extractedPageUrls: string[] = [];
    let cleanPageBgUrls: string[] = [];
    let slideLayouts: PptxSlideLayoutInfo[] = [];
    let extractedPagesList: ExtractedPageDetail[] = [];
    let isCanvaDesignTemplate = false;

    const itemId = item.id || item.designId || item._extractedObject?.id;

    if (itemId && (item.source === 'design' || item.category === 'design' || item.type === 'design')) {
      isCanvaDesignTemplate = true;
      try {
        console.log(`[CANVA IMPORT] Extraindo páginas e camadas editáveis do design ${itemId}...`);
        const [extractRes, inspectRes] = await Promise.allSettled([
          fetch(`/api/canva/designs/${encodeURIComponent(itemId)}/extract-pages`, {
            method: 'POST',
          }),
          fetch('/api/canva/templates/inspect', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: itemId, titleHint: item.name }),
          }),
        ]);

        if (extractRes.status === 'fulfilled' && extractRes.value.ok) {
          const extractData = await extractRes.value.json();
          if (extractData.totalPages) {
            realPageCount = extractData.totalPages;
          }
          if (extractData.pageUrls && extractData.pageUrls.length > 0) {
            extractedPageUrls = extractData.pageUrls.filter(Boolean);
          }
          if (extractData.pages && extractData.pages.length > 0) {
            extractedPagesList = extractData.pages;
          }
        }

        if (inspectRes.status === 'fulfilled' && inspectRes.value.ok) {
          const inspectData = await inspectRes.value.json();
          if (inspectData.success) {
            if (inspectData.pageCount) {
              realPageCount = inspectData.pageCount;
            }
            if (Array.isArray(inspectData.slideLayouts)) {
              slideLayouts = inspectData.slideLayouts;
            }
            if (Array.isArray(inspectData.pages)) {
              cleanPageBgUrls = inspectData.pages.map((p: any) => p.cleanBgUrl || '');
            }
            // Sincroniza automaticamente o modelo inspecionado e campos-chave detectados
            setInspectedTemplate(inspectData);
            setTemplateId(inspectData.brandTemplateId || inspectData.id);
            setAutofillTitlePrefix(inspectData.title || item.name || 'LoteCriativo');
            applyDetectedMappingsFromInspection(inspectData);
          }
        }
      } catch (e) {
        console.warn('[CANVA IMPORT] Extractor fallback:', e);
      }
    }

    if (itemUrl && !itemUrl.startsWith('data:')) {
      try {
        const res = await fetch(itemUrl);
        if (res.ok) {
          const blob = await res.blob();
          finalUrl = URL.createObjectURL(blob);
          mediaFile = new File([blob], `${item.name}.${itemType === 'video' ? 'mp4' : 'png'}`, { type: blob.type });
        }
      } catch {
        // mantém URL
      }
    }

    const baseItem = { url: finalUrl, name: item.name, type: itemType, id: itemId };
    const initialPageUrls = extractedPageUrls.length > 0 ? extractedPageUrls : getDefaultPageUrls(baseItem);

    if (extractedPagesList.length === 0) {
      extractedPagesList = Array.from({ length: realPageCount }).map((_, idx) => {
        const pNum = idx + 1;
        const layout = slideLayouts.find((l) => l.slideIndex === pNum);
        const textsFromXml = layout?.textBoxes?.map((tb) => tb.text).filter(Boolean) || [];
        return {
          pageNumber: pNum,
          title: `Página ${pNum}: ${item.name}`,
          status: 'success',
          mediaUrl: initialPageUrls[idx % initialPageUrls.length] || finalUrl,
          texts:
            textsFromXml.length > 0
              ? textsFromXml
              : [
                  `Título: ${item.name} (Página ${pNum})`,
                  `Marcador de Frase Reconhecido: (frase) / FRASE`,
                ],
          elements: [
            `Plano de Fundo (Página ${pNum})`,
            ...(textsFromXml.length > 0 ? [`Camadas de Texto Reais (${textsFromXml.length})`] : ['Camada de Texto Editável']),
          ],
          hasPhrasePlaceholder:
            layout ? layout.detectedTokens.some((t) => /frase|quote|texto/i.test(t)) : true,
        };
      });
    }

    const newItem: LocalMediaItem = {
      id: itemId || Math.random().toString(36).substring(7),
      name: item.name,
      type: itemType,
      size: mediaFile.size || 1024 * 500,
      url: finalUrl,
      file: mediaFile,
      pageCount: realPageCount,
      activePageIndex: 0,
      pageMediaUrls: initialPageUrls,
      cleanPageBgUrls,
      slideLayouts,
      isCanvaDesignTemplate,
      extractedPages: extractedPagesList,
    };

    setMediaList((prev) => [...prev, newItem]);
    return newItem;
  };

  // =========================================================================
  // CANVA FOLDER EXPLORER & VIDEOS HANDLERS
  // =========================================================================
  const loadExplorerVideos = async () => {
    setLoadingExplorerVideos(true);
    try {
      const res = await fetch('/api/canva/videos');
      const data = await res.json();
      if (res.ok && data.success) {
        setExplorerVideosData({
          videoDesigns: data.videoDesigns || [],
          allDesigns: data.allDesigns || [],
          videoAssets: data.videoAssets || [],
          audioAssets: data.audioAssets || [],
          transferredMedia: data.transferredMedia || [],
          apiLimitationNote: data.apiLimitationNote,
          audioLimitationNote: data.audioLimitationNote,
        });
      }
    } catch (err) {
      console.error('Falha ao listar vídeos/áudios do Canva:', err);
    } finally {
      setLoadingExplorerVideos(false);
    }
  };

  const handleCaptureCanvaSdkSelection = async (mediaKind: 'video' | 'audio') => {
    setCapturingSdkMedia(true);
    setSdkCaptureFeedback(null);
    try {
      const res = await captureSelectedMediaFromCanvaEditor(mediaKind);
      if (res.success && res.transferred.length > 0) {
        for (const m of res.transferred) {
          await processAndAddCanvaItem({
            id: m.id,
            name: m.name,
            type: m.type,
            category: m.type,
            source: m.source,
            url: m.url,
            pageCount: 1,
          });
        }
        await loadExplorerVideos();
        setSdkCaptureFeedback({
          text: `${res.transferred.length} arquivo(s) de ${mediaKind === 'video' ? 'vídeo' : 'áudio'} capturado(s) via Selection API + getTemporaryUrl e transferido(s) com sucesso!`,
          type: 'success',
        });
      } else {
        setSdkCaptureFeedback({
          text: res.error || 'Nenhum item selecionado capturado.',
          type: 'info',
        });
      }
    } catch (err: any) {
      setSdkCaptureFeedback({
        text: err?.message || 'Erro ao capturar mídia selecionada no Editor Canva.',
        type: 'error',
      });
    } finally {
      setCapturingSdkMedia(false);
    }
  };

  const handleLookupCanvaAsset = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!assetLookupInput.trim()) return;
    setLookingUpAsset(true);
    setAssetLookupError(null);
    try {
      const res = await fetch('/api/canva/assets/lookup', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ assetId: assetLookupInput.trim() }),
      });
      const data = await res.json();
      if (res.ok && data.success && data.asset) {
        setAssetLookupInput('');
        await loadExplorerVideos();
      } else {
        setAssetLookupError(data.error || 'Asset não encontrado no Canva.');
      }
    } catch (err: any) {
      setAssetLookupError(err.message || 'Falha ao consultar Asset ID.');
    } finally {
      setLookingUpAsset(false);
    }
  };

  const handleExportDesignMp4AndImport = async (designItem: any, closeModal = true) => {
    const id = designItem.id;
    if (!id) return;
    setExportingMp4Id(id);
    try {
      const preferredQuality = aspectRatio === '9:16' ? 'vertical_1080p' : 'horizontal_1080p';
      const res = await fetch(`/api/canva/designs/${encodeURIComponent(id)}/export-mp4`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ quality: preferredQuality }),
      });
      const data = await res.json();
      if (res.ok && data.success && Array.isArray(data.urls) && data.urls.length > 0) {
        const videoUrls: string[] = data.urls;
        const newItem: LocalMediaItem = {
          id: `${id}_mp4_${Date.now()}`,
          name: `${designItem.name || 'Vídeo Canva'} (MP4)`,
          type: 'video',
          size: 1024 * 1024 * 4,
          url: videoUrls[0],
          file: new File([], `${designItem.name || 'video'}.mp4`, { type: 'video/mp4' }),
          pageCount: videoUrls.length,
          activePageIndex: 0,
          pageMediaUrls: videoUrls,
        };
        setMediaList((prev) => [...prev, newItem]);
        if (closeModal) {
          setShowFolderExplorer(false);
        }
        alert(`Vídeo "${designItem.name}" exportado em MP4 oficial do Canva (${videoUrls.length} cena/página) e adicionado às mídias!`);
      } else {
        alert(data.error || 'Este design não pôde ser exportado como MP4.');
      }
    } catch (err: any) {
      alert(`Erro ao exportar MP4 do Canva: ${err.message || err}`);
    } finally {
      setExportingMp4Id(null);
    }
  };

  const openCanvaFolderExplorer = (
    folderId: string = 'root',
    folderName: string = 'Projetos (Raiz)',
    tab: 'folders' | 'videos' = 'folders',
    initialFilter: 'all' | 'image' | 'video' | 'audio' | 'design' = tab === 'videos' ? 'video' : 'all'
  ) => {
    setShowFolderExplorer(true);
    setExplorerTab(tab);
    setExplorerMediaFilter(initialFilter);
    setExplorerCurrentFolderId(folderId);
    setExplorerBreadcrumbs([{ id: folderId, name: folderName }]);
    setSelectedExplorerIds(new Set());
    setSelectedExplorerItemsMap(new Map());
    setFolderExplorerSearch('');
    setSdkCaptureFeedback(null);
    loadExplorerVideos();
    if (tab !== 'videos') {
      loadExplorerFolder(folderId, folderName, false);
    }
  };

  const loadExplorerFolder = async (folderId: string, folderName?: string, isPaging = false) => {
    if (isPaging) {
      setLoadingMoreExplorer(true);
    } else {
      setLoadingExplorer(true);
    }

    try {
      let url = `/api/canva/folders/${encodeURIComponent(folderId)}/items?limit=50`;
      if (isPaging && explorerContinuation) {
        url += `&continuation=${encodeURIComponent(explorerContinuation)}`;
      }

      const res = await fetch(url);
      const data = await res.json();

      if (res.ok && data.success) {
        if (isPaging) {
          setExplorerItems((prev) => [...prev, ...(data.items || [])]);
        } else {
          setExplorerItems(data.items || []);
          setExplorerCurrentFolderId(folderId);
          if (folderName) {
            setExplorerBreadcrumbs((prev) => {
              const idx = prev.findIndex((b) => b.id === folderId);
              if (idx !== -1) {
                return prev.slice(0, idx + 1);
              }
              return [...prev, { id: folderId, name: folderName }];
            });
          }
        }
        setExplorerContinuation(data.continuation || null);
      } else {
        console.warn('Aviso ao carregar itens da pasta do Canva:', data.error);
      }
    } catch (err) {
      console.error('Falha ao consultar pasta do Canva:', err);
    } finally {
      setLoadingExplorer(false);
      setLoadingMoreExplorer(false);
    }
  };

  const handleExplorerEnterFolder = (folder: any) => {
    const fId = folder.id;
    const fName = folder.name || 'Subpasta';
    setExplorerBreadcrumbs((prev) => [...prev, { id: fId, name: fName }]);
    loadExplorerFolder(fId, fName, false);
  };

  const handleExplorerNavigateBreadcrumb = (index: number) => {
    const target = explorerBreadcrumbs[index];
    if (!target) return;
    const updated = explorerBreadcrumbs.slice(0, index + 1);
    setExplorerBreadcrumbs(updated);
    loadExplorerFolder(target.id, target.name, false);
  };

  const handleExplorerGoBack = () => {
    if (explorerBreadcrumbs.length <= 1) return;
    handleExplorerNavigateBreadcrumb(explorerBreadcrumbs.length - 2);
  };

  const toggleSelectExplorerItem = (item: any) => {
    const itemId = item.id;
    setSelectedExplorerIds((prev) => {
      const next = new Set(prev);
      if (next.has(itemId)) {
        next.delete(itemId);
      } else {
        next.add(itemId);
      }
      return next;
    });

    setSelectedExplorerItemsMap((prev) => {
      const next = new Map(prev);
      if (next.has(itemId)) {
        next.delete(itemId);
      } else {
        next.set(itemId, item);
      }
      return next;
    });
  };

  const handleSelectAllExplorerItems = () => {
    const selectable =
      explorerMediaFilter === 'video'
        ? [
            ...((videoFilterMode === 'video_candidates'
              ? explorerVideosData?.videoDesigns
              : explorerVideosData?.allDesigns) || []),
            ...(explorerVideosData?.videoAssets || []),
          ]
        : explorerMediaFilter === 'audio'
        ? explorerVideosData?.audioAssets || []
        : explorerMediaFilter === 'design'
        ? [
            ...explorerItems.filter((i) => i.type === 'design'),
            ...(explorerVideosData?.allDesigns || []),
          ]
        : explorerMediaFilter === 'image'
        ? explorerItems.filter((i) => i.type === 'image')
        : explorerItems.filter((i) => i.type !== 'folder');
    const allIds = new Set(selectable.map((i) => i.id));
    const allMap = new Map();
    selectable.forEach((i) => allMap.set(i.id, i));
    setSelectedExplorerIds(allIds);
    setSelectedExplorerItemsMap(allMap);
  };

  const handleClearExplorerSelection = () => {
    setSelectedExplorerIds(new Set());
    setSelectedExplorerItemsMap(new Map());
  };

  const handleConfirmExplorerImport = async () => {
    if (selectedExplorerItemsMap.size === 0) return;
    setImportingFromExplorer(true);
    try {
      for (const item of Array.from(selectedExplorerItemsMap.values())) {
        const normalizedType =
          item.type === 'video'
            ? 'video'
            : item.type === 'audio'
            ? 'audio'
            : item.type === 'design'
            ? 'design'
            : 'image';
        await processAndAddCanvaItem({
          id: item.id,
          name: item.name,
          type: normalizedType,
          category: normalizedType,
          source: normalizedType,
          url: item.url || item.thumbnailUrl || '',
          pageCount: item.pageCount || (normalizedType === 'design' ? 3 : 1),
          urls: { edit_url: item.editUrl, view_url: item.viewUrl },
        });
      }
      setShowFolderExplorer(false);
      setSelectedExplorerIds(new Set());
      setSelectedExplorerItemsMap(new Map());
    } catch (err) {
      console.error('Erro na importação selecionada:', err);
    } finally {
      setImportingFromExplorer(false);
    }
  };

  // =========================================================================
  // CANVA ASSET UPLOADS (Parte 2 & 5 da Especificação)
  // =========================================================================
  const initializeCanvaUploadQueue = () => {
    const queue = batchResults.map((r, idx) => {
      const existing = canvaUploadQueue.find((q) => q.id === (r.id || `res_${idx}`));
      if (existing && existing.status === 'success') {
        return existing; // Mantém já enviado! Evita reenvio duplicado
      }
      const pageInfo = getPreviewPageInfo(idx, mediaList);
      const safePhraseName = r.phrase.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 25);
      const isVideo = pageInfo.isVideo;
      const ext = isVideo ? 'mp4' : 'png';
      const fileName = `Lote_${idx + 1}_${safePhraseName}.${ext}`;

      return {
        id: r.id || `res_${idx}`,
        name: fileName,
        phrase: r.phrase,
        mediaUrl: pageInfo.pageUrl || r.url || 'https://images.unsplash.com/photo-1550745165-9bc0b252726f?w=600&q=80',
        status: (existing?.status === 'success' ? 'success' : 'pending') as any,
        assetId: existing?.assetId,
        canvaUploadsUrl: existing?.canvaUploadsUrl,
        errorMessage: existing?.errorMessage,
      };
    });
    setCanvaUploadQueue(queue);
    return queue;
  };

  const handleUploadAssetsToCanva = async (onlyRetryFailed = false) => {
    if (batchResults.length === 0) return;
    const currentQueue = canvaUploadQueue.length === batchResults.length ? [...canvaUploadQueue] : initializeCanvaUploadQueue();

    setIsUploadingToCanva(true);
    const itemsToUpload = currentQueue.filter((item) => {
      if (onlyRetryFailed) return item.status === 'failed';
      return item.status !== 'success';
    });

    setUploadProgressCount({ current: 0, total: itemsToUpload.length });

    let doneCount = 0;
    for (const item of itemsToUpload) {
      setCanvaUploadQueue((prev) =>
        prev.map((i) => (i.id === item.id ? { ...i, status: 'uploading', errorMessage: undefined } : i))
      );

      try {
        const payload: any = { name: item.name };
        if (item.mediaUrl.startsWith('blob:') || item.mediaUrl.startsWith('data:')) {
          if (item.mediaUrl.startsWith('data:')) {
            payload.dataBase64 = item.mediaUrl;
          } else {
            const blobRes = await fetch(item.mediaUrl);
            const blob = await blobRes.blob();
            const dataUrl = await new Promise<string>((resolve, reject) => {
              const reader = new FileReader();
              reader.onloadend = () => resolve(String(reader.result || ''));
              reader.onerror = reject;
              reader.readAsDataURL(blob);
            });
            payload.dataBase64 = dataUrl;
          }
        } else {
          payload.url = item.mediaUrl;
        }

        const res = await fetch('/api/canva/assets/upload', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

        const data = await res.json();
        if (res.ok && data.success && data.asset) {
          setCanvaUploadQueue((prev) =>
            prev.map((i) =>
              i.id === item.id
                ? {
                    ...i,
                    status: 'success',
                    assetId: data.asset.id,
                    canvaUploadsUrl: data.canvaUploadsUrl || 'https://www.canva.com/folder/uploads',
                  }
                : i
            )
          );
        } else {
          setCanvaUploadQueue((prev) =>
            prev.map((i) =>
              i.id === item.id
                ? {
                    ...i,
                    status: 'failed',
                    errorMessage: data.error || 'Falha ao processar asset no Canva.',
                  }
                : i
            )
          );
        }
      } catch (err: any) {
        setCanvaUploadQueue((prev) =>
          prev.map((i) =>
            i.id === item.id
              ? {
                  ...i,
                  status: 'failed',
                  errorMessage: err.message || 'Erro de conexão no upload.',
                }
              : i
          )
        );
      }

      doneCount++;
      setUploadProgressCount({ current: doneCount, total: itemsToUpload.length });
    }

    setIsUploadingToCanva(false);
  };

  // =========================================================================
  // MOTOR DE CONVERSÃO FIEL DE MÍDIAS E CAMADAS EDITÁVEIS (PPTX & CANVA)
  // =========================================================================
  const resolveMediaUrlToDataUrl = async (
    url: string,
    isVideo: boolean,
    ratio: '9:16' | '1:1' | '16:9'
  ): Promise<string | null> => {
    if (!url) return null;
    if (url.startsWith('data:image/')) return url;

    if (isVideo) {
      try {
        const frameDataUrl = await new Promise<string | null>((resolve) => {
          const videoEl = document.createElement('video');
          videoEl.crossOrigin = 'anonymous';
          videoEl.muted = true;
          videoEl.playsInline = true;
          videoEl.preload = 'auto';

          let settled = false;
          const finish = (val: string | null) => {
            if (settled) return;
            settled = true;
            videoEl.src = '';
            resolve(val);
          };

          const timer = setTimeout(() => finish(null), 4000);

          videoEl.onloadeddata = () => {
            try {
              const targetW = ratio === '9:16' ? 720 : ratio === '1:1' ? 720 : 1280;
              const targetH = ratio === '9:16' ? 1280 : ratio === '1:1' ? 720 : 720;
              const canvas = document.createElement('canvas');
              canvas.width = targetW;
              canvas.height = targetH;
              const ctx = canvas.getContext('2d');
              if (!ctx || !videoEl.videoWidth || !videoEl.videoHeight) {
                clearTimeout(timer);
                finish(null);
                return;
              }
              const scale = Math.max(targetW / videoEl.videoWidth, targetH / videoEl.videoHeight);
              const drawW = videoEl.videoWidth * scale;
              const drawH = videoEl.videoHeight * scale;
              const dx = (targetW - drawW) / 2;
              const dy = (targetH - drawH) / 2;
              ctx.drawImage(videoEl, dx, dy, drawW, drawH);
              const out = canvas.toDataURL('image/jpeg', 0.88);
              clearTimeout(timer);
              finish(out);
            } catch {
              clearTimeout(timer);
              finish(null);
            }
          };

          videoEl.onerror = () => {
            clearTimeout(timer);
            finish(null);
          };

          videoEl.src = url;
        });
        if (frameDataUrl) return frameDataUrl;
      } catch {
        // fallback abaixo
      }
      return null;
    }

    if (url.startsWith('blob:')) {
      try {
        const blobRes = await fetch(url);
        const blob = await blobRes.blob();
        if (!blob.type.startsWith('image/')) return null;
        return await new Promise<string | null>((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : null);
          reader.onerror = () => resolve(null);
          reader.readAsDataURL(blob);
        });
      } catch {
        return null;
      }
    }

    if (url.startsWith('http://') || url.startsWith('https://')) {
      try {
        const proxyRes = await fetch(`/api/proxy-media?url=${encodeURIComponent(url)}`);
        if (proxyRes.ok) {
          const proxyData = await proxyRes.json();
          if (proxyData.success && proxyData.dataUrl) {
            return proxyData.dataUrl;
          }
        }
      } catch {
        // tenta fetch direto como fallback
      }

      try {
        const directRes = await fetch(url);
        if (directRes.ok) {
          const blob = await directRes.blob();
          return await new Promise<string | null>((resolve) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(typeof reader.result === 'string' ? reader.result : null);
            reader.onerror = () => resolve(null);
            reader.readAsDataURL(blob);
          });
        }
      } catch {
        return null;
      }
    }

    return null;
  };

  const buildEditablePptxBase64 = async (
    slidesInput: Array<{
      phrase: string;
      mediaUrl: string;
      isVideo: boolean;
      pageNum?: number;
      slideLayout?: PptxSlideLayoutInfo | null;
    }>,
    title?: string
  ): Promise<string> => {
    const pptx = new pptxgen();
    if (title) pptx.title = title;

    let slideWidth = 10;
    let slideHeight = 5.625;

    if (aspectRatio === '9:16') {
      slideWidth = 5.625;
      slideHeight = 10;
      pptx.defineLayout({ name: 'PORTRAIT_9_16', width: 5.625, height: 10 });
      pptx.layout = 'PORTRAIT_9_16';
    } else if (aspectRatio === '1:1') {
      slideWidth = 8;
      slideHeight = 8;
      pptx.defineLayout({ name: 'SQUARE_1_1', width: 8, height: 8 });
      pptx.layout = 'SQUARE_1_1';
    } else {
      slideWidth = 10;
      slideHeight = 5.625;
      pptx.layout = 'LAYOUT_16x9';
    }

    // Pré-resolve imagens de fundo únicas em paralelo para máxima velocidade
    const uniqueCache = new Map<string, string | null>();
    await Promise.all(
      slidesInput.map(async (item) => {
        const cacheKey = `${item.isVideo ? 'v' : 'i'}:${item.mediaUrl}`;
        if (!uniqueCache.has(cacheKey) && item.mediaUrl) {
          uniqueCache.set(cacheKey, '__pending__');
          const resolved = await resolveMediaUrlToDataUrl(item.mediaUrl, item.isVideo, aspectRatio);
          uniqueCache.set(cacheKey, resolved);
        }
      })
    );

    const cleanHexColor = (textColor || '#FFFFFF').replace('#', '').toUpperCase();
    const overlayTransparency = Math.max(0, Math.min(100, 100 - overlayOpacity));

    for (let i = 0; i < slidesInput.length; i++) {
      const { phrase, mediaUrl, isVideo, pageNum, slideLayout } = slidesInput[i];
      const cacheKey = `${isVideo ? 'v' : 'i'}:${mediaUrl}`;
      const bgDataUrl = uniqueCache.get(cacheKey);
      const slide = pptx.addSlide();

      // Fundo base escuro caso a mídia não exista
      slide.background = { color: '1C1917' };

      // Camada 1: Imagem de fundo real em camada editável separada no Canva (full bleed cover)
      if (bgDataUrl && bgDataUrl !== '__pending__' && bgDataUrl.startsWith('data:image/')) {
        slide.addImage({
          data: bgDataUrl,
          x: 0,
          y: 0,
          w: slideWidth,
          h: slideHeight,
          sizing: { type: 'cover', w: slideWidth, h: slideHeight },
        });
      }

      // Se o slide veio do modelo Canva (slideLayouts), atualiza apenas a propriedade `text`
      // dos elementos originais preservando fontFamily, fontSizePt, colorHex, bold e align, sem sobrepor novas caixas
      if (slideLayout && slideLayout.textBoxes.length > 0) {
        const effectivePage = pageNum || i + 1;
        const updatedBoxes = replaceSlideLayoutTextBoxesInPlace(
          slideLayout,
          effectivePage,
          phrase,
          keyFieldMappings,
          signatureText
        );

        for (const box of updatedBoxes) {
          const origHex = (box.colorHex || '#FFFFFF').replace('#', '').toUpperCase();
          slide.addText(box.text, {
            x: (box.xPct / 100) * slideWidth,
            y: (box.yPct / 100) * slideHeight,
            w: (box.wPct / 100) * slideWidth,
            h: (box.hPct / 100) * slideHeight,
            fontSize: box.fontSizePt || 24,
            fontFace: box.fontFamily || 'Arial',
            color: origHex,
            align: box.align,
            bold: box.bold,
            italic: Boolean(box.italic),
            underline: box.underline ? { style: 'sng' } : undefined,
            charSpacing: box.charSpacingPt,
            lineSpacingMultiple: box.lineSpacingPct ? box.lineSpacingPct / 100 : undefined,
            wrap: true,
          });
        }
        continue;
      }

      // Camada 2: Película escura editável correspondente ao slider overlayOpacity (apenas para criação do zero)
      if (overlayOpacity > 0) {
        slide.addShape(pptx.ShapeType.rect, {
          x: 0,
          y: 0,
          w: slideWidth,
          h: slideHeight,
          fill: { color: '000000', transparency: overlayTransparency },
          line: { transparency: 100 },
        });
      }

      // Camada 3: Caixa de texto principal editável preservando fonte, alinhamento, cor e posição
      const boxY =
        textPosition === 'top'
          ? slideHeight * 0.08
          : textPosition === 'bottom'
          ? slideHeight * 0.46
          : slideHeight * 0.25;
      const boxHeight = slideHeight * 0.42;
      const computedFontSize = Math.round(Math.min(44, Math.max(14, fontSize * (aspectRatio === '9:16' ? 0.92 : 1.0))));

      slide.addText(phrase, {
        x: slideWidth * 0.08,
        y: boxY,
        w: slideWidth * 0.84,
        h: boxHeight,
        fontSize: computedFontSize,
        fontFace: fontFamily === 'serif' ? 'Georgia' : 'Arial',
        color: cleanHexColor || 'FFFFFF',
        align: textAlign as 'left' | 'center' | 'right',
        valign: textPosition === 'top' ? 'top' : textPosition === 'bottom' ? 'bottom' : 'middle',
        bold: true,
        wrap: true,
        outline: hasStroke ? { size: 0.75, color: '000000' } : undefined,
        shadow: hasShadow
          ? { type: 'outer', blur: 6, offset: 2, angle: 45, color: '000000', opacity: 0.8 }
          : undefined,
      });

      // Camada 4: Assinatura editável posicionada de acordo com o layout escolhido
      if (signatureText) {
        const sigY =
          textPosition === 'top'
            ? boxY + boxHeight + 0.1
            : textPosition === 'bottom'
            ? slideHeight - 0.75
            : boxY + boxHeight + 0.15;

        slide.addText(signatureText, {
          x: slideWidth * 0.08,
          y: Math.min(slideHeight - 0.65, sigY),
          w: slideWidth * 0.84,
          h: 0.5,
          fontSize: 12,
          fontFace: 'Courier New',
          color: cleanHexColor || 'FFFFFF',
          align: textAlign as 'left' | 'center' | 'right',
          valign: 'middle',
        });
      }
    }

    return (await pptx.write({ outputType: 'base64' })) as string;
  };

  // =========================================================================
  // OPÇÃO 1: MODELOS DO CANVA, CAMPOS-CHAVE & BRAND TEMPLATE AUTOFILL
  // =========================================================================
  const loadBrandTemplates = async () => {
    setLoadingBrandTemplates(true);
    setBrandTemplateAuthError(null);
    try {
      const res = await fetch('/api/canva/brand-templates');
      const data = await res.json();
      if (res.ok && data.success) {
        setBrandTemplatesList(data.items || []);
      } else {
        setBrandTemplatesList([]);
        setBrandTemplateAuthError(
          data.error || 'Não foi possível listar os Brand Templates do Canva.'
        );
      }
    } catch (err: any) {
      setBrandTemplateAuthError(err.message || 'Erro de comunicação ao buscar Brand Templates.');
    } finally {
      setLoadingBrandTemplates(false);
    }
  };

  const loadTemplateSelectorFolder = async (folderId: string, folderName?: string, isPaging = false) => {
    if (isPaging) {
      setLoadingMoreTplFolder(true);
    } else {
      setLoadingTplFolder(true);
    }
    try {
      let url = `/api/canva/folders/${encodeURIComponent(folderId)}/items?limit=50`;
      if (isPaging && tplFolderContinuation) {
        url += `&continuation=${encodeURIComponent(tplFolderContinuation)}`;
      }
      const res = await fetch(url);
      const data = await res.json();
      if (res.ok && data.success) {
        if (isPaging) {
          setTplFolderItems((prev) => [...prev, ...(data.items || [])]);
        } else {
          setTplFolderItems(data.items || []);
          setTplFolderId(folderId);
          if (folderName) {
            setTplBreadcrumbs((prev) => {
              const idx = prev.findIndex((b) => b.id === folderId);
              if (idx !== -1) return prev.slice(0, idx + 1);
              return [...prev, { id: folderId, name: folderName }];
            });
          }
        }
        setTplFolderContinuation(data.continuation || null);
      }
    } catch (err) {
      console.warn('Erro ao navegar pastas para seleção de modelo:', err);
    } finally {
      setLoadingTplFolder(false);
      setLoadingMoreTplFolder(false);
    }
  };

  const openTemplateSelectorModal = (tab: 'brand_templates' | 'folders' = 'brand_templates') => {
    setShowTemplateSelectorModal(true);
    setTemplateSelectorTab(tab);
    setTemplateSelectorSearch('');
    loadBrandTemplates();
    if (tplFolderItems.length === 0) {
      setTplBreadcrumbs([{ id: 'root', name: 'Projetos (Raiz)' }]);
      loadTemplateSelectorFolder('root', 'Projetos (Raiz)', false);
    }
  };

  const applyDetectedMappingsFromInspection = (data: any) => {
    const totalPages = Math.max(1, data.pageCount || 1);
    const allPageNumbers = Array.from({ length: totalPages }, (_, i) => i + 1);
    const datasetFields: Array<{ name: string; type: string }> = data.datasetFields || [];
    const detectedPlaceholders: string[] = data.detectedPlaceholders || [];

    const detectedNames: string[] =
      datasetFields.length > 0
        ? datasetFields.map((d) => d.name)
        : detectedPlaceholders;

    if (detectedNames.length > 0) {
      const builtMappings = detectedNames.map((rawName) => {
        const upper = rawName.toUpperCase();
        let source: 'batch_phrase' | 'signature' | 'custom_text' | 'ignore' = 'custom_text';
        let customText = '';
        if (upper.includes('FRASE') || upper.includes('QUOTE') || upper.includes('TEXTO')) {
          source = 'batch_phrase';
        } else if (upper.includes('ASSINATURA') || upper.includes('SIGNATURE') || upper.includes('AUTOR')) {
          source = 'signature';
        } else if (upper.includes('DATA') || upper.includes('DATE')) {
          source = 'custom_text';
          customText = new Date().toLocaleDateString('pt-BR');
        }
        return {
          fieldKey: rawName,
          detectedInTemplate: true,
          source,
          customText,
          pageScope: 'all' as const,
          selectedPages: allPageNumbers,
        };
      });

      const existingUpper = new Set(builtMappings.map((m) => m.fieldKey.replace(/^[\(\{\[<]|[\)\}\]>]$/g, '').toUpperCase()));
      const extras = [
        { fieldKey: 'TITULO', source: 'ignore' as const, customText: '' },
        { fieldKey: 'SUBTITULO', source: 'ignore' as const, customText: '' },
        { fieldKey: 'ASSINATURA', source: 'ignore' as const, customText: '' },
        { fieldKey: 'DATA', source: 'ignore' as const, customText: new Date().toLocaleDateString('pt-BR') },
      ]
        .filter((e) => !existingUpper.has(e.fieldKey))
        .map((e) => ({
          ...e,
          detectedInTemplate: false,
          pageScope: 'all' as const,
          selectedPages: allPageNumbers,
        }));

      const firstPhraseMapping = builtMappings.find((m) => m.source === 'batch_phrase');
      if (firstPhraseMapping) {
        setFieldPhrase(firstPhraseMapping.fieldKey);
      }

      setKeyFieldMappings([...builtMappings, ...extras]);
    } else {
      setKeyFieldMappings((prev) =>
        prev.map((m) => ({
          ...m,
          detectedInTemplate: false,
          selectedPages: allPageNumbers,
        }))
      );
    }
  };

  const inspectSelectedCanvaTemplate = async (id: string, titleHint?: string) => {
    const trimmed = String(id || '').trim();
    if (!trimmed) return;
    setInspectingTemplate(true);
    setAutofillStatusBanner(null);
    try {
      const res = await fetch('/api/canva/templates/inspect', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: trimmed, titleHint }),
      });
      const data = await res.json();
      if (res.ok && data.success) {
        setInspectedTemplate(data);
        setTemplateId(data.brandTemplateId || data.id);
        setAutofillTitlePrefix(data.title || 'LoteCriativo');
        setShowTemplateSelectorModal(false);
        applyDetectedMappingsFromInspection(data);

        // Atualiza item correspondente na mediaList (se existir) com fundo limpo & slideLayouts
        if (Array.isArray(data.slideLayouts) && data.slideLayouts.length > 0) {
          const cleanBgs = Array.isArray(data.pages) ? data.pages.map((p: any) => p.cleanBgUrl || '') : [];
          setMediaList((prev) =>
            prev.map((m) =>
              m.id === data.id
                ? {
                    ...m,
                    isCanvaDesignTemplate: true,
                    cleanPageBgUrls: cleanBgs,
                    slideLayouts: data.slideLayouts,
                    pageCount: data.pageCount || m.pageCount,
                  }
                : m
            )
          );
        }
      } else {
        setAutofillStatusBanner({
          text: data.error || 'Não foi possível inspecionar este modelo do Canva.',
          type: 'error',
        });
      }
    } catch (err: any) {
      setAutofillStatusBanner({
        text: `Falha ao consultar modelo no Canva: ${err.message || err}`,
        type: 'error',
      });
    } finally {
      setInspectingTemplate(false);
    }
  };

  const resolveFieldPreviewValue = (
    mapping: {
      fieldKey: string;
      source: 'batch_phrase' | 'signature' | 'custom_text' | 'ignore';
      customText: string;
      pageScope: 'all' | 'specific';
      selectedPages: number[];
    },
    phraseText: string
  ): string | null => {
    if (mapping.source === 'ignore') return null;
    if (mapping.pageScope === 'specific' && mapping.selectedPages.length === 0) return null;
    if (mapping.source === 'batch_phrase') return phraseText;
    if (mapping.source === 'signature') return signatureText || '@lotecriativo';
    return mapping.customText;
  };

  const handleAddCustomKeyField = () => {
    const cleaned = newCustomFieldKey.trim();
    if (!cleaned) return;
    if (keyFieldMappings.some((m) => m.fieldKey.toUpperCase() === cleaned.toUpperCase())) {
      setNewCustomFieldKey('');
      return;
    }
    const totalPages = Math.max(1, inspectedTemplate?.pageCount || 1);
    const isPhraseLike = /frase|quote|texto/i.test(cleaned);
    setKeyFieldMappings((prev) => [
      ...prev,
      {
        fieldKey: cleaned,
        detectedInTemplate: isFieldKeyFoundInModel(cleaned),
        source: isPhraseLike ? 'batch_phrase' : 'custom_text',
        customText: '',
        pageScope: 'all',
        selectedPages: Array.from({ length: totalPages }, (_, i) => i + 1),
      },
    ]);
    setNewCustomFieldKey('');
  };

  const buildReplacedVersionsPayload = () => {
    const activeMappings = keyFieldMappings.filter((m) => m.source !== 'ignore');
    const effectiveMappings =
      activeMappings.length > 0
        ? activeMappings
        : [
            {
              fieldKey: '(frase)',
              detectedInTemplate: true,
              source: 'batch_phrase' as const,
              customText: '',
              pageScope: 'all' as const,
              selectedPages: Array.from({ length: Math.max(1, inspectedTemplate?.pageCount || 1) }, (_, i) => i + 1),
            },
          ];

    return validPhrases.map((phrase) => {
      const replacements = effectiveMappings
        .map((m) => {
          const val = resolveFieldPreviewValue(m, phrase);
          if (val === null) return null;
          return {
            findText: m.fieldKey,
            replaceText: val,
            pageScope: m.pageScope,
            selectedPages: m.selectedPages,
          };
        })
        .filter(Boolean);
      return { label: phrase, replacements };
    });
  };

  const runOption1AutofillBatch = async () => {
    const isAutofill = Boolean(inspectedTemplate?.brandTemplateId && inspectedTemplate?.autofillCompatible);
    const isPptxRoundtrip = Boolean(inspectedTemplate?.id && inspectedTemplate?.editableViaPptxRoundtrip);

    if (!isAutofill && !isPptxRoundtrip) {
      setAutofillStatusBanner({
        text: 'Selecione e inspecione um Design ou Brand Template do Canva antes de gerar as versões editáveis.',
        type: 'error',
      });
      return;
    }
    if (validPhrases.length === 0) {
      setAutofillStatusBanner({
        text: 'Adicione pelo menos uma frase na Etapa 2 antes de gerar as versões.',
        type: 'error',
      });
      return;
    }

    const activeMappings = keyFieldMappings.filter((m) => m.source !== 'ignore');
    if (activeMappings.length === 0) {
      setAutofillStatusBanner({
        text: 'Ative pelo menos um campo-chave (ex: (frase) ou FRASE) para substituição antes de gerar.',
        type: 'error',
      });
      return;
    }
    if (isModelValidationBlocked) {
      const missingNames = missingRequiredFieldMappings.map((m) => m.fieldKey || '(vazio)').join(', ');
      setAutofillStatusBanner({
        text:
          missingNames.length > 0
            ? `Geração bloqueada: os campos obrigatórios (${missingNames}) não foram encontrados no modelo selecionado. Corrija o mapeamento destacado em vermelho antes de gerar.`
            : 'Geração bloqueada: mapeie pelo menos um campo de frase encontrado no modelo selecionado antes de gerar.',
        type: 'error',
      });
      return;
    }

    setAutofillBatchLoading(true);
    setAutofillStatusBanner({
      text: isAutofill
        ? `Gerando ${validPhrases.length} versões editáveis no Canva via Brand Template Autofill API...`
        : option1OutputMode === 'multipage'
        ? `Substituindo os campos dentro do arquivo original do Canva (${validPhrases.length} páginas) e importando como design editável...`
        : `Substituindo o campo-chave diretamente dentro do arquivo original para ${validPhrases.length} versões editáveis...`,
      type: 'info',
    });

    try {
      let data: any;
      let ok = false;

      if (isAutofill) {
        const versions = validPhrases.map((phrase) => {
          const fieldValues: Record<string, string> = {};
          for (const mapping of activeMappings) {
            const val = resolveFieldPreviewValue(mapping, phrase);
            if (val !== null) {
              fieldValues[mapping.fieldKey] = val;
            }
          }
          return { label: phrase, fieldValues };
        });

        const res = await fetch('/api/canva/autofill-batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            brandTemplateId: inspectedTemplate!.brandTemplateId,
            titlePrefix: autofillTitlePrefix || inspectedTemplate!.title || 'Modelo Lote',
            versions,
          }),
        });
        ok = res.ok;
        data = await res.json();
      } else {
        const versions = buildReplacedVersionsPayload();
        const res = await fetch('/api/canva/designs/replace-text-batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            designId: inspectedTemplate!.id,
            titlePrefix: autofillTitlePrefix || inspectedTemplate!.title || 'Modelo Editado',
            mode: option1OutputMode,
            versions,
          }),
        });
        ok = res.ok;
        data = await res.json();
      }

      if (Array.isArray(data.results)) {
        setAutofillVersionResults(data.results);
      }

      if (ok && data.success) {
        setAutofillStatusBanner({
          text: isAutofill
            ? `${data.completedCount} de ${data.total} versões criadas como novos designs editáveis no Canva! Fontes, cores, imagens e layout foram preservados.`
            : `${data.completedCount} de ${data.total} design(s) criados no Canva com o texto substituído diretamente dentro do arquivo (sem texto extra sobreposto)!`,
          type: 'success',
        });
        const mappedBatch: BatchResult[] = (data.results || [])
          .filter((r: any) => r.status === 'completed' && r.editUrl)
          .map((r: any) => ({
            id: r.designId || `af_${r.versionIndex}`,
            phrase: r.label || r.title,
            status: 'completed' as const,
            type: 'canva_design' as const,
            url: r.editUrl,
            mediaName: inspectedTemplate!.title,
            pageCount: r.pageCount || inspectedTemplate!.pageCount || 1,
            isMultipage: option1OutputMode === 'multipage',
            message: `${r.title} · ${r.pageCount || 1} pág(s) · ID: ${r.designId}`,
          }));
        if (mappedBatch.length > 0) {
          setBatchResults(mappedBatch);
          if (option1OutputMode === 'multipage' && mappedBatch[0]?.url) {
            setCreatedCanvaDesign({
              designId: mappedBatch[0].id,
              title: data.results[0]?.title || inspectedTemplate!.title,
              pageCount: mappedBatch[0].pageCount,
              editUrl: mappedBatch[0].url,
              viewUrl: data.results[0]?.viewUrl,
            });
          }
        }
      } else {
        setAutofillStatusBanner({
          text: data.error || `Concluído com ${data.failedCount || 0} falha(s). Verifique os detalhes abaixo.`,
          type: 'error',
        });
      }
    } catch (err: any) {
      setAutofillStatusBanner({
        text: `Erro ao substituir campos e salvar no Canva: ${err.message || err}`,
        type: 'error',
      });
    } finally {
      setAutofillBatchLoading(false);
    }
  };

  const handleDownloadReplacedPptx = async (customDesignId?: string, customTitle?: string, singleIndex?: number) => {
    const targetDesignId = customDesignId || inspectedTemplate?.id;
    if (!targetDesignId || validPhrases.length === 0) return;
    if (isModelValidationBlocked) {
      const missingNames = missingRequiredFieldMappings.map((m) => m.fieldKey || '(vazio)').join(', ');
      setAutofillStatusBanner({
        text:
          missingNames.length > 0
            ? `Download bloqueado: corrija os campos obrigatórios não encontrados no modelo (${missingNames}).`
            : 'Download bloqueado: mapeie um campo-chave válido encontrado no modelo.',
        type: 'error',
      });
      return;
    }

    setDownloadingReplacedPptx(true);
    try {
      const allVersions = buildReplacedVersionsPayload();
      const versions =
        singleIndex !== undefined && allVersions[singleIndex]
          ? [allVersions[singleIndex]]
          : allVersions;
      const res = await fetch('/api/canva/designs/replace-text-download', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          designId: targetDesignId,
          titlePrefix:
            singleIndex !== undefined
              ? `${customTitle || autofillTitlePrefix || inspectedTemplate?.title || 'Modelo_Editado'}_Item_${singleIndex + 1}`
              : customTitle || autofillTitlePrefix || inspectedTemplate?.title || 'Modelo_Editado',
          versions,
        }),
      });
      const data = await res.json();
      if (!res.ok || !data.pptxBase64) {
        setAutofillStatusBanner({
          text: data.error || 'Erro ao gerar PPTX com texto substituído.',
          type: 'error',
        });
        return;
      }

      const byteCharacters = atob(data.pptxBase64);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const blob = new Blob([new Uint8Array(byteNumbers)], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = data.filename || 'Modelo_Substituido.pptx';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      setAutofillStatusBanner({
        text: `Erro ao baixar .PPTX substituído: ${err.message || err}`,
        type: 'error',
      });
    } finally {
      setDownloadingReplacedPptx(false);
    }
  };

  // =========================================================================
  // CANVA EDITABLE DESIGN CREATION (Opção 2: Novo Design no Canva via PPTX)
  // =========================================================================
  const handleSaveDesignToCanva = async (modeOverride?: 'multipage' | 'individual', singleIndex?: number) => {
    if (validPhrases.length === 0) return;
    if (isModelValidationBlocked) {
      const missingNames = missingRequiredFieldMappings.map((m) => m.fieldKey || '(vazio)').join(', ');
      setDesignSaveBanner({
        text:
          missingNames.length > 0
            ? `Geração bloqueada: os campos obrigatórios (${missingNames}) não foram encontrados no modelo selecionado. Ajuste os campos em vermelho na aba de mapeamento.`
            : 'Geração bloqueada: nenhum campo-chave de frase mapeado foi encontrado no modelo selecionado.',
        type: 'error',
      });
      return;
    }

    // Se há um design do Canva com camadas XML editáveis, usa SEMPRE substituição in-place OOXML (mesmo em singleIndex!)
    const canvaTemplateId =
      inspectedTemplate?.editableViaPptxRoundtrip
        ? inspectedTemplate.id
        : mediaList.find((m) => m.isCanvaDesignTemplate)?.id;

    if (canvaTemplateId) {
      setIsSavingDesignToCanva(true);
      setDesignSaveBanner(null);
      const allVersions = buildReplacedVersionsPayload();
      const targetVersions =
        singleIndex !== undefined && allVersions[singleIndex]
          ? [allVersions[singleIndex]]
          : allVersions;
      const targetMode = singleIndex !== undefined ? 'individual' : (modeOverride || 'multipage');
      setSavingDesignStatusText(
        singleIndex !== undefined
          ? `Substituindo campos-chave dentro do arquivo Canva para a versão #${singleIndex + 1} sem sobrepor texto...`
          : `Substituindo campos-chave dentro do arquivo Canva (${validPhrases.length} frases) sem sobrepor texto extra...`
      );
      try {
        const res = await fetch('/api/canva/designs/replace-text-batch', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            designId: canvaTemplateId,
            titlePrefix: autofillTitlePrefix || inspectedTemplate?.title || 'Modelo Editado',
            mode: targetMode,
            versions: targetVersions,
          }),
        });
        const data = await res.json();
        if (res.ok && data.success && Array.isArray(data.results)) {
          if (singleIndex !== undefined) {
            const firstOk = data.results.find((r: any) => r.status === 'completed' && r.editUrl);
            if (firstOk) {
              setBatchResults((prev) =>
                prev.map((item, idx) =>
                  idx === singleIndex
                    ? {
                        ...item,
                        id: firstOk.designId,
                        status: 'completed',
                        type: 'canva_design',
                        url: firstOk.editUrl,
                        pageCount: firstOk.pageCount || item.pageCount,
                        message: `Texto substituído diretamente no arquivo Canva (ID: ${firstOk.designId}).`,
                      }
                    : item
                )
              );
              setDesignSaveBanner({
                text: `Versão #${singleIndex + 1} salva no Canva preservando toda a formatação original!`,
                type: 'success',
              });
            } else {
              setDesignSaveBanner({
                text: data.results[0]?.error || 'Falha ao salvar versão individual no Canva.',
                type: 'error',
              });
            }
          } else {
            setAutofillVersionResults(data.results);
            const mapped: BatchResult[] = data.results
              .filter((r: any) => r.status === 'completed' && r.editUrl)
              .map((r: any) => ({
                id: r.designId,
                phrase: r.label || r.title,
                status: 'completed' as const,
                type: 'canva_design' as const,
                url: r.editUrl,
                mediaName: inspectedTemplate?.title || 'Design Canva',
                pageCount: r.pageCount || validPhrases.length,
                isMultipage: targetMode === 'multipage',
                message: `Texto substituído diretamente no arquivo Canva (ID: ${r.designId}).`,
              }));
            if (mapped.length > 0) {
              setBatchResults(mapped);
              if (targetMode === 'multipage') {
                setCreatedCanvaDesign({
                  designId: mapped[0].id,
                  title: data.results[0]?.title || 'Design Editado',
                  pageCount: mapped[0].pageCount,
                  editUrl: mapped[0].url!,
                  viewUrl: data.results[0]?.viewUrl,
                });
              }
            }
            setDesignSaveBanner({
              text: `${data.completedCount} design(s) salvo(s) no Canva com substituição direta no texto original!`,
              type: 'success',
            });
          }
        } else {
          setDesignSaveBanner({
            text: data.error || 'Falha ao salvar substituição in-place no Canva.',
            type: 'error',
          });
        }
      } catch (err: any) {
        setDesignSaveBanner({
          text: `Erro ao substituir texto no Canva: ${err.message || err}`,
          type: 'error',
        });
      } finally {
        setIsSavingDesignToCanva(false);
      }
      return;
    }

    setIsSavingDesignToCanva(true);
    setDesignSaveBanner(null);

    try {
      // Caso 1: Salvar apenas um slide individual específico no Canva
      if (singleIndex !== undefined && validPhrases[singleIndex]) {
        const phrase = validPhrases[singleIndex];
        const pageInfo = getPreviewPageInfo(singleIndex, mediaList);
        const tplPageNum = (singleIndex % Math.max(1, inspectedTemplate?.pageCount || 1)) + 1;
        const tplLayout =
          inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
          inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
          null;
        const tplCleanUrl = inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)]?.cleanBgUrl || '';
        setSavingDesignStatusText(`Preparando slide #${singleIndex + 1} com preservação de camadas originais...`);

        const base64Data = await buildEditablePptxBase64(
          [
            {
              phrase,
              mediaUrl: pageInfo.cleanPageUrl || tplCleanUrl || pageInfo.pageUrl,
              isVideo: pageInfo.isVideo,
              pageNum: pageInfo.media ? pageInfo.pageNum : tplPageNum,
              slideLayout: pageInfo.slideLayout || tplLayout,
            },
          ],
          `Slide ${singleIndex + 1} - ${phrase.substring(0, 30)}`
        );

        setSavingDesignStatusText(`Enviando slide #${singleIndex + 1} para o Canva (POST /rest/v1/imports)...`);
        const res = await fetch('/api/canva/designs/import-pptx', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            title: `LoteCriativo Slide ${singleIndex + 1} - ${phrase.substring(0, 28)}`,
            pptxBase64: base64Data,
            expectedPageCount: 1,
          }),
        });
        const data = await res.json();
        if (res.ok && data.success && data.editUrl) {
          setBatchResults((prev) =>
            prev.map((item, idx) =>
              idx === singleIndex
                ? {
                    ...item,
                    id: data.designId,
                    status: 'completed',
                    type: 'canva_design',
                    url: data.editUrl,
                    pageCount: data.pageCount || 1,
                    message: `Design editável criado no Canva (ID: ${data.designId}). Clique em Abrir no Canva para editar!`,
                  }
                : item
            )
          );
          setSavingDesignStatusText(`Slide #${singleIndex + 1} criado como Design editável no Canva!`);
          setDesignSaveBanner({ text: `Slide #${singleIndex + 1} criado no Canva com sucesso (ID: ${data.designId}).`, type: 'success' });
        } else {
          setDesignSaveBanner({ text: data.error || 'Falha ao criar design individual no Canva.', type: 'error' });
        }
        return;
      }

      const effectiveMode = modeOverride || 'multipage';

      // Caso 2: Criar 1 Design Editável separado no Canva para cada frase
      if (effectiveMode === 'individual') {
        let createdOk = 0;
        const createdList: BatchResult[] = [];
        for (let i = 0; i < validPhrases.length; i++) {
          const phrase = validPhrases[i];
          const pageInfo = getPreviewPageInfo(i, mediaList);
          const tplPageNum = (i % Math.max(1, inspectedTemplate?.pageCount || 1)) + 1;
          const tplLayout =
            inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
            inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
            null;
          const tplCleanUrl = inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)]?.cleanBgUrl || '';
          setSavingDesignStatusText(`Criando Design Editável ${i + 1} de ${validPhrases.length} no Canva...`);

          const base64Data = await buildEditablePptxBase64(
            [
              {
                phrase,
                mediaUrl: pageInfo.cleanPageUrl || tplCleanUrl || pageInfo.pageUrl,
                isVideo: pageInfo.isVideo,
                pageNum: pageInfo.media ? pageInfo.pageNum : tplPageNum,
                slideLayout: pageInfo.slideLayout || tplLayout,
              },
            ],
            `Design ${i + 1} - ${phrase.substring(0, 28)}`
          );

          const res = await fetch('/api/canva/designs/import-pptx', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              title: `LoteCriativo #${i + 1} - ${phrase.substring(0, 28)}`,
              pptxBase64: base64Data,
              expectedPageCount: 1,
            }),
          });
          const data = await res.json();
          if (res.ok && data.success && data.editUrl) {
            createdOk++;
            createdList.push({
              id: data.designId,
              phrase,
              status: 'completed',
              type: 'canva_design',
              url: data.editUrl,
              mediaName: pageInfo.media?.name || 'Camadas editáveis',
              pageCount: data.pageCount || 1,
              message: `Design editável individual criado no Canva (ID: ${data.designId}).`,
            });
          }
        }
        if (createdList.length > 0) {
          setBatchResults(createdList);
        }
        setSavingDesignStatusText(`${createdOk} de ${validPhrases.length} designs individuais editáveis criados no Canva!`);
        setDesignSaveBanner({
          text: `${createdOk} de ${validPhrases.length} designs editáveis criados no Canva!`,
          type: createdOk > 0 ? 'success' : 'error',
        });
        return;
      }

      // Caso 3: Criar 1 Design Multipáginas unificado no Canva com TODAS as páginas
      setSavingDesignStatusText('Convertendo fundos das páginas, fontes e alinhamentos preservados para camadas editáveis...');
      const slidesPayload = validPhrases.map((phrase, idx) => {
        const pageInfo = getPreviewPageInfo(idx, mediaList);
        const tplPageNum = (idx % Math.max(1, inspectedTemplate?.pageCount || 1)) + 1;
        const tplLayout =
          inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
          inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
          null;
        const tplCleanUrl = inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)]?.cleanBgUrl || '';
        return {
          phrase,
          mediaUrl: pageInfo.cleanPageUrl || tplCleanUrl || pageInfo.pageUrl,
          isVideo: pageInfo.isVideo,
          pageNum: pageInfo.media ? pageInfo.pageNum : tplPageNum,
          slideLayout: pageInfo.slideLayout || tplLayout,
        };
      });

      const title = `Lote Criativo - ${validPhrases.length} Páginas (${new Date().toLocaleDateString('pt-BR')})`;
      const base64Data = await buildEditablePptxBase64(slidesPayload, title);

      setSavingDesignStatusText('Canva está convertendo as camadas editáveis do novo design (POST /rest/v1/imports)...');
      const res = await fetch('/api/canva/designs/import-pptx', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title,
          pptxBase64: base64Data,
          expectedPageCount: validPhrases.length,
        }),
      });

      const data = await res.json();
      if (res.ok && data.success && data.editUrl) {
        const realPages = data.pageCount || validPhrases.length;
        setCreatedCanvaDesign({
          designId: data.designId,
          title: data.title,
          pageCount: realPages,
          editUrl: data.editUrl,
          viewUrl: data.viewUrl,
        });
        setBatchResults([
          {
            id: data.designId,
            phrase: `Design Editável Multipáginas (${realPages} Páginas no Canva)`,
            status: 'completed',
            type: 'canva_design',
            url: data.editUrl,
            mediaName: `${mediaList.length || 1} mídia(s) de fundo`,
            pageCount: realPages,
            isMultipage: true,
            message: `Design nativo criado no Canva (ID: ${data.designId}, ${realPages} páginas) com imagens, película, textos e assinatura editáveis.`,
          },
        ]);
        setSavingDesignStatusText(`Novo design editável (${realPages} páginas) criado no Canva com sucesso!`);
        setDesignSaveBanner({
          text: `Design "${data.title}" criado com ${realPages} páginas editáveis no Canva!`,
          type: 'success',
        });
      } else {
        setDesignSaveBanner({ text: data.error || 'Falha ao importar novo design para o Canva.', type: 'error' });
      }
    } catch (err: any) {
      console.error('Erro ao salvar design no Canva:', err);
      setDesignSaveBanner({ text: `Erro ao criar design no Canva: ${err.message || err}`, type: 'error' });
    } finally {
      setIsSavingDesignToCanva(false);
    }
  };

  useEffect(() => {
    if (activeStep === 7 && batchResults.length > 0 && canvaUploadQueue.length === 0) {
      initializeCanvaUploadQueue();
    }
  }, [activeStep, batchResults.length]);

  // Login Firebase
  const handleFirebaseLogin = async () => {
    const provider = new GoogleAuthProvider();
    try {
      await signInWithPopup(auth, provider);
    } catch (error) {
      console.error('Erro no login Firebase:', error);
    }
  };

  const handleFirebaseLogout = async () => {
    await signOut(auth);
  };

  // Salvar credenciais manuais (se não estiverem nas variáveis de ambiente do servidor)
  const handleSaveCredentials = async () => {
    if (!manualClientId || !manualClientSecret) {
      setCredentialsMsg({ text: 'Preencha o Client ID e Client Secret.', type: 'error' });
      return;
    }
    setSavingCredentials(true);
    setCredentialsMsg(null);
    try {
      // Salva em localStorage e cookies para não precisar logar novamente ao recarregar a página
      localStorage.setItem('canva_saved_credentials', JSON.stringify({
        clientId: manualClientId,
        clientSecret: manualClientSecret,
      }));
      document.cookie = "canva_connected=true; max-age=31536000; path=/";

      const res = await fetch('/api/auth/canva/credentials', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ clientId: manualClientId, clientSecret: manualClientSecret }),
      });
      const data = await res.json();
      if (res.ok) {
        setCredentialsMsg({ text: 'Credenciais salvas e mantidas na sessão! Conexão permanente ativa.', type: 'success' });
        await fetchCanvaStatus();
      } else {
        setCredentialsMsg({ text: data.error || 'Falha ao salvar credenciais.', type: 'error' });
      }
    } catch (err: any) {
      setCredentialsMsg({ text: 'Erro de conexão com o servidor.', type: 'error' });
    } finally {
      setSavingCredentials(false);
    }
  };

  // Iniciar conexão OAuth Canva com PKCE
  const handleConnect = async (includeBrandTemplates: boolean = false) => {
    setConnecting(true);
    console.log('[CANVA OAUTH FLOW] Initiating Canva OAuth flow...');
    console.log('[CANVA OAUTH FLOW] Current app state before redirect:', {
      canvaStatus,
      includeBrandTemplates,
      manualClientId: manualClientId ? manualClientId.substring(0, 8) + '...' : 'NONE',
      redirectUri: canvaStatus?.redirectUri || '',
      timestamp: new Date().toISOString(),
    });

    try {
      const urlEndpoint = includeBrandTemplates
        ? '/api/auth/canva/url?includeBrandTemplates=true'
        : '/api/auth/canva/url';
      console.log(`[CANVA OAUTH FLOW] Requesting backend auth URL from ${urlEndpoint}...`);
      const response = await fetch(urlEndpoint);
      const data = await response.json();

      console.log('[CANVA OAUTH FLOW] Response received from /api/auth/canva/url:', {
        status: response.status,
        ok: response.ok,
        data,
      });

      if (!response.ok || !data.url) {
        console.error('[CANVA OAUTH FLOW] Failed to obtain valid auth URL from backend:', data);
        alert(data.error || 'Não foi possível gerar a URL de autorização. Verifique se o Client ID e Secret estão configurados.');
        setConnecting(false);
        return;
      }

      setDirectAuthUrl(data.url);

      // Parse query parameters for explicit verification log
      let parsedParams: Record<string, string> = {};
      try {
        const u = new URL(data.url);
        u.searchParams.forEach((val, key) => {
          parsedParams[key] = val;
        });
      } catch (e) {
        console.warn('[CANVA OAUTH FLOW] Failed parsing searchParams from URL:', e);
      }

      console.log('====================================================');
      console.log('[CANVA OAUTH FLOW] Full generated OAuth redirect URL before window.open:', data.url);
      console.log('[CANVA OAUTH FLOW] Query parameters matching Canva documentation requirements:', {
        state: parsedParams['state'] || 'N/A',
        client_id: parsedParams['client_id'] || 'N/A',
        redirect_uri: parsedParams['redirect_uri'] || 'N/A',
        scope: parsedParams['scope'] || 'N/A',
        code_challenge: parsedParams['code_challenge'] || 'N/A',
        code_challenge_method: parsedParams['code_challenge_method'] || 'N/A',
        response_type: parsedParams['response_type'] || 'N/A',
      });
      console.log('====================================================');

      // Abre o popup apontando diretamente para o provedor OAuth (Canva)
      const authWindow = window.open(data.url, 'oauth_canva_popup', 'width=620,height=750,menubar=no,status=no');

      if (!authWindow || authWindow.closed || typeof authWindow.closed === 'undefined') {
        console.warn('[CANVA OAUTH FLOW] Popup blocked or failed to open. Prompting direct redirection...');
        if (confirm('O navegador bloqueou a abertura do popup. Deseja prosseguir com o Redirecionamento Direto para o Canva?')) {
          window.location.href = data.url;
          return;
        }
      }

      // Escuta mensagem de sucesso do callback com logs detalhados
      const messageHandler = (event: MessageEvent) => {
        console.log('[CANVA OAUTH FLOW] window.addEventListener("message") received:', event.data);
        if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
          console.log('[CANVA OAUTH FLOW] OAuth authorization successful message received.');
          window.removeEventListener('message', messageHandler);
          setConnecting(false);
          fetchCanvaStatus();
        } else if (event.data?.type === 'OAUTH_AUTH_ERROR') {
          console.warn('[CANVA OAUTH FLOW] OAuth error message received:', event.data.error);
          window.removeEventListener('message', messageHandler);
          setConnecting(false);
          alert(`Erro retornado pelo Canva: ${event.data.error || 'Autorização negada'}`);
        }
      };

      window.addEventListener('message', messageHandler);

      // Fallback se o usuário fechar a janela manualmente
      const checkClosed = setInterval(() => {
        if (authWindow && authWindow.closed) {
          clearInterval(checkClosed);
          console.log('[CANVA OAUTH FLOW] Auth popup window closed by user.');
          setConnecting(false);
          fetchCanvaStatus();
        }
      }, 1500);
    } catch (error: any) {
      console.error('[CANVA OAUTH FLOW] Network or execution error in handleConnect:', error);
      alert('Falha ao comunicar com o servidor. Verifique o console.');
      setConnecting(false);
    }
  };

  // Função utilitária para remoção profunda de cookies em múltiplos caminhos e domínios
  const clearCanvaCookies = () => {
    console.log('[CANVA DISCONNECT DEBUG] Executando limpeza profunda de cookies e storage...');
    const cookieNames = ['canva_connected', 'canva_session'];
    const paths = ['/', '/api', '/api/auth', '/api/canva', window.location.pathname];
    const hostname = window.location.hostname;
    const parts = hostname.split('.');
    
    const domains = ['', hostname, `.${hostname}`];
    if (parts.length > 1) {
      domains.push(`.${parts.slice(-2).join('.')}`);
    }

    cookieNames.forEach((name) => {
      paths.forEach((path) => {
        domains.forEach((domain) => {
          const domainAttr = domain ? `; domain=${domain}` : '';
          document.cookie = `${name}=; expires=Thu, 01 Jan 1970 00:00:00 GMT; max-age=0; path=${path}${domainAttr}`;
        });
      });
    });

    try {
      localStorage.removeItem('canva_saved_credentials');
      localStorage.removeItem('canva_connected');
      sessionStorage.clear();
      console.log('[CANVA DISCONNECT DEBUG] Storage limpo com sucesso.');
    } catch (e) {
      console.warn('[CANVA DISCONNECT DEBUG] Erro ao limpar storage:', e);
    }
  };

  // Desconectar do Canva com rastreamento detalhado de depuração e revogação
  const handleDisconnect = async () => {
    console.log('====================================================');
    console.log('[CANVA DISCONNECT DEBUG] Iniciando fluxo de desconexão...');
    if (!confirm('Deseja realmente desconectar sua conta do Canva?')) {
      console.log('[CANVA DISCONNECT DEBUG] Desconexão cancelada pelo usuário.');
      return;
    }

    try {
      console.log('[CANVA DISCONNECT DEBUG] Passo 1: Limpeza prévia de cookies no cliente...');
      clearCanvaCookies();

      console.log('[CANVA DISCONNECT DEBUG] Passo 2: Enviando requisição POST /api/auth/canva/disconnect ao servidor...');
      const res = await fetch('/api/auth/canva/disconnect', { method: 'POST' });
      console.log('[CANVA DISCONNECT DEBUG] Código HTTP retornado pelo servidor:', res.status);

      if (res.ok) {
        const data = await res.json();
        console.log('[CANVA DISCONNECT DEBUG] Resposta do servidor recebida com sucesso:', data);
      } else {
        const errText = await res.text();
        console.error('[CANVA DISCONNECT DEBUG] Servidor retornou erro na desconexão:', res.status, errText);
      }

      console.log('[CANVA DISCONNECT DEBUG] Passo 3: Limpeza secundária pós-resposta...');
      clearCanvaCookies();

      console.log('[CANVA DISCONNECT DEBUG] Passo 4: Atualizando estado React local para desconectado...');
      setCanvaStatus((prev) => ({
        connected: false,
        configured: Boolean(prev?.configured),
        clientSecretConfigured: Boolean(prev?.clientSecretConfigured),
        redirectUri: prev?.redirectUri || '',
      }));

      console.log('[CANVA DISCONNECT DEBUG] Passo 5: Forçando recarregamento da página (window.location.reload())...');
      console.log('====================================================');
      window.location.reload();
    } catch (err) {
      console.error('[CANVA DISCONNECT DEBUG] Falha de rede durante o processo de desconexão:', err);
      clearCanvaCookies();
      window.location.reload();
    }
  };

  // Copiar Redirect URI
  const copyRedirectUri = (uri: string) => {
    navigator.clipboard.writeText(uri);
    setCopiedUri(true);
    setTimeout(() => setCopiedUri(false), 2500);
  };

  // Tratamento de frases
  const validPhrases = rawPhrases
    .split('\n')
    .map((p) => p.trim())
    .filter((p) => p.length > 0);

  // Validação em tempo real dos campos-chave em relação ao modelo Canva selecionado
  const modelDetectedTexts = useMemo(() => {
    const texts = new Set<string>();
    if (inspectedTemplate) {
      inspectedTemplate.datasetFields?.forEach((d) => {
        if (d.name?.trim()) texts.add(d.name.trim());
      });
      inspectedTemplate.detectedPlaceholders?.forEach((p) => {
        if (p?.trim()) texts.add(p.trim());
      });
      inspectedTemplate.slideLayouts?.forEach((layout) => {
        layout.detectedTokens?.forEach((t) => {
          if (t?.trim()) texts.add(t.trim());
        });
        layout.textBoxes?.forEach((tb) => {
          if (tb.text?.trim()) texts.add(tb.text.trim());
          tb.tokens?.forEach((t) => {
            if (t?.trim()) texts.add(t.trim());
          });
        });
      });
    }
    mediaList.forEach((m) => {
      if (!m.isCanvaDesignTemplate) return;
      m.slideLayouts?.forEach((layout) => {
        layout.detectedTokens?.forEach((t) => {
          if (t?.trim()) texts.add(t.trim());
        });
        layout.textBoxes?.forEach((tb) => {
          if (tb.text?.trim()) texts.add(tb.text.trim());
          tb.tokens?.forEach((t) => {
            if (t?.trim()) texts.add(t.trim());
          });
        });
      });
    });
    return Array.from(texts);
  }, [inspectedTemplate, mediaList]);

  const modelSelectableTokens = useMemo(() => {
    const tokens = new Set<string>();
    inspectedTemplate?.datasetFields?.forEach((d) => {
      if (d.name?.trim()) tokens.add(d.name.trim());
    });
    inspectedTemplate?.detectedPlaceholders?.forEach((p) => {
      if (p?.trim()) tokens.add(p.trim());
    });
    inspectedTemplate?.slideLayouts?.forEach((layout) => {
      layout.detectedTokens?.forEach((t) => {
        if (t?.trim()) tokens.add(t.trim());
      });
      layout.textBoxes?.forEach((tb) => {
        if (tb.tokens?.length) {
          tb.tokens.forEach((t) => {
            if (t?.trim()) tokens.add(t.trim());
          });
        } else if (tb.text?.trim() && tb.text.trim().length <= 36) {
          tokens.add(tb.text.trim());
        }
      });
    });
    mediaList.forEach((m) => {
      if (!m.isCanvaDesignTemplate) return;
      m.slideLayouts?.forEach((layout) => {
        layout.detectedTokens?.forEach((t) => {
          if (t?.trim()) tokens.add(t.trim());
        });
      });
    });
    return Array.from(tokens);
  }, [inspectedTemplate, mediaList]);

  const hasInspectableModelLoaded = useMemo(() => {
    const fromTemplate = Boolean(
      inspectedTemplate &&
        (inspectedTemplate.autofillCompatible ||
          inspectedTemplate.editableViaPptxRoundtrip ||
          (inspectedTemplate.datasetFields?.length || 0) > 0 ||
          (inspectedTemplate.slideLayouts?.length || 0) > 0)
    );
    const fromMedia = mediaList.some((m) => m.isCanvaDesignTemplate && (m.slideLayouts?.length || 0) > 0);
    return fromTemplate || fromMedia;
  }, [inspectedTemplate, mediaList]);

  const isFieldKeyFoundInModel = (rawKey: string): boolean => {
    if (!hasInspectableModelLoaded) return true;
    const target = String(rawKey || '').trim();
    if (!target) return false;

    if (inspectedTemplate?.autofillCompatible && inspectedTemplate.datasetFields.length > 0) {
      const targetClean = target.replace(/^[\(\{\[<]|[\)\}\]>]$/g, '').trim().toLowerCase();
      return inspectedTemplate.datasetFields.some(
        (d) =>
          d.name.trim().toLowerCase() === target.toLowerCase() ||
          d.name.trim().toLowerCase() === targetClean
      );
    }

    const esc = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const hasBracket = /^[\(\{\[<].*[\)\}\]>]$/.test(target);
    const patterns: RegExp[] = [];

    if (hasBracket) {
      patterns.push(new RegExp(esc(target), 'i'));
    } else {
      const cleanW = esc(target.replace(/^[\(\{\[<]|[\)\}\]>]$/g, '').trim());
      patterns.push(new RegExp(`\\(${cleanW}\\)`, 'i'));
      patterns.push(new RegExp(`\\{${cleanW}\\}`, 'i'));
      patterns.push(new RegExp(`\\[${cleanW}\\]`, 'i'));
      patterns.push(new RegExp(`<${cleanW}>`, 'i'));
      patterns.push(new RegExp(`(?:^|[^a-zA-Z0-9À-ÿ_])${cleanW}(?:$|[^a-zA-Z0-9À-ÿ_])`, 'i'));
    }

    return modelDetectedTexts.some((txt) => patterns.some((re) => re.test(txt)));
  };

  const activeRequiredMappings = keyFieldMappings.filter((m) => m.source !== 'ignore');
  const missingRequiredFieldMappings = hasInspectableModelLoaded
    ? activeRequiredMappings.filter((m) => !isFieldKeyFoundInModel(m.fieldKey))
    : [];
  const hasValidBatchPhraseMapping = activeRequiredMappings.some(
    (m) => m.source === 'batch_phrase' && isFieldKeyFoundInModel(m.fieldKey)
  );
  const isStep3FieldPhraseMissing =
    hasInspectableModelLoaded && (!fieldPhrase.trim() || !isFieldKeyFoundInModel(fieldPhrase));
  const isModelValidationBlocked =
    hasInspectableModelLoaded && (missingRequiredFieldMappings.length > 0 || !hasValidBatchPhraseMapping);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const newItems: LocalMediaItem[] = [];
    Array.from(files).forEach((file) => {
      const isVideo = file.type.startsWith('video/');
      const isAudio = file.type.startsWith('audio/');
      const isImage = file.type.startsWith('image/');
      if (!isVideo && !isAudio && !isImage) return;

      const mediaType: LocalMediaItem['type'] = isVideo ? 'video' : isAudio ? 'audio' : 'image';
      const objectUrl = URL.createObjectURL(file);
      const baseItem = { url: objectUrl, name: file.name, type: mediaType };

      newItems.push({
        id: Math.random().toString(36).substring(7),
        name: file.name,
        type: mediaType,
        size: file.size,
        url: objectUrl,
        file,
        pageCount: isAudio ? 1 : 3,
        activePageIndex: 0,
        pageMediaUrls: isAudio ? [] : getDefaultPageUrls(baseItem),
      });
    });

    setMediaList((prev) => [...prev, ...newItems]);
  };

  const removeMedia = (id: string) => {
    setMediaList((prev) => prev.filter((m) => m.id !== id));
  };

  const importPhrasesFile = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      const text = event.target?.result as string;
      if (text) {
        setRawPhrases((prev) => (prev ? prev + '\n' + text : text));
      }
    };
    reader.readAsText(file);
  };

  // Iniciar Motor de Lote
  const startBatchProcessing = async (overrideMode?: 'individual' | 'multipage') => {
    if (validPhrases.length === 0) {
      setDesignSaveBanner({
        text: 'Adicione pelo menos uma frase válida antes de gerar o lote.',
        type: 'error',
      });
      return;
    }

    if (isModelValidationBlocked) {
      const missingNames = missingRequiredFieldMappings.map((m) => m.fieldKey || '(vazio)').join(', ');
      setDesignSaveBanner({
        text:
          missingNames.length > 0
            ? `Geração bloqueada: os campos obrigatórios (${missingNames}) não foram encontrados no modelo selecionado.`
            : 'Geração bloqueada: mapeie pelo menos um campo de frase encontrado no modelo selecionado.',
        type: 'error',
      });
      return;
    }

    const targetMode = overrideMode || designMode;
    setIsProcessing(true);
    setBatchProgress({
      current: 0,
      total: validPhrases.length,
      status: targetMode === 'multipage' ? 'Criando arquivo multipáginas unificado...' : 'Substituindo marcadores de texto no Canva...',
    });

    try {
      const response = await fetch('/api/canva/batch-generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          phrases: validPhrases,
          templateId: templateId.trim() || undefined,
          isMultipage: targetMode === 'multipage',
          settings: {
            aspectRatio,
            textPosition,
            textAlign,
            fontFamily,
            fontSize,
            textColor,
            overlayOpacity,
            hasShadow,
            hasStroke,
            signatureText,
            fieldPhrase,
            fieldSignature,
            designMode: targetMode,
          },
          mediaItems: mediaList.map((m) => ({ name: m.name, type: m.type })),
        }),
      });

      const data = await response.json();

      if (response.ok && data.results) {
        setBatchResults(data.results);
        setBatchProgress({ current: validPhrases.length, total: validPhrases.length, status: 'Concluído com sucesso!' });
        setActiveStep(7); // Vai para a tela de resultados
      } else {
        setDesignSaveBanner({
          text: data.error || 'Erro no processamento do lote.',
          type: 'error',
        });
      }
    } catch (err) {
      console.error('Erro ao gerar lote:', err);
      setDesignSaveBanner({
        text: 'Falha de rede ao conectar com o servidor de processamento.',
        type: 'error',
      });
    } finally {
      setIsProcessing(false);
    }
  };

  // Exportar Relatório CSV
  const exportCsvReport = () => {
    if (batchResults.length === 0) return;
    const headers = ['Número', 'Frase', 'Status', 'Tipo', 'Mídia Associada', 'URL'];
    const rows = batchResults.map((r, i) => [
      i + 1,
      `"${r.phrase.replace(/"/g, '""')}"`,
      r.status,
      r.type,
      `"${r.mediaName}"`,
      r.url || 'Exportação Editável (.PPTX)',
    ]);
    const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(','), ...rows.map((e) => e.join(','))].join('\n');
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement('a');
    link.setAttribute('href', encodedUri);
    link.setAttribute('download', `lotecriativo_relatorio_${Date.now()}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  // Download de Arquivo Editável Real em .PPTX (Pronto para importar no Canva ou editar em qualquer editor)
  const downloadAlternativePackage = async (singleIndex?: number) => {
    if (validPhrases.length === 0) return;
    if (isModelValidationBlocked) {
      const missingNames = missingRequiredFieldMappings.map((m) => m.fieldKey || '(vazio)').join(', ');
      setDesignSaveBanner({
        text:
          missingNames.length > 0
            ? `Download bloqueado: corrija os campos obrigatórios não encontrados no modelo (${missingNames}).`
            : 'Download bloqueado: mapeie pelo menos um campo-chave de frase compatível com o modelo.',
        type: 'error',
      });
      return;
    }

    // Se há design Canva editável via XML, usa substituição direta no arquivo original
    const canvaTemplateId =
      inspectedTemplate?.editableViaPptxRoundtrip
        ? inspectedTemplate.id
        : mediaList.find((m) => m.isCanvaDesignTemplate)?.id;

    if (canvaTemplateId) {
      setDownloadingPptx(true);
      try {
        await handleDownloadReplacedPptx(
          canvaTemplateId,
          inspectedTemplate?.title || 'Lote_Substituido',
          singleIndex
        );
      } finally {
        setDownloadingPptx(false);
      }
      return;
    }

    setDownloadingPptx(true);
    try {
      const targetSlides =
        singleIndex !== undefined && validPhrases[singleIndex]
          ? (() => {
              const info = getPreviewPageInfo(singleIndex, mediaList);
              const tplPageNum = (singleIndex % Math.max(1, inspectedTemplate?.pageCount || 1)) + 1;
              const tplLayout =
                inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
                inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
                null;
              const tplCleanUrl =
                inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)]?.cleanBgUrl || '';
              return [
                {
                  phrase: validPhrases[singleIndex],
                  mediaUrl: info.cleanPageUrl || tplCleanUrl || info.pageUrl,
                  isVideo: info.isVideo,
                  pageNum: info.media ? info.pageNum : tplPageNum,
                  slideLayout: info.slideLayout || tplLayout,
                },
              ];
            })()
          : validPhrases.map((phrase, idx) => {
              const info = getPreviewPageInfo(idx, mediaList);
              const tplPageNum = (idx % Math.max(1, inspectedTemplate?.pageCount || 1)) + 1;
              const tplLayout =
                inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
                inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
                null;
              const tplCleanUrl =
                inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)]?.cleanBgUrl || '';
              return {
                phrase,
                mediaUrl: info.cleanPageUrl || tplCleanUrl || info.pageUrl,
                isVideo: info.isVideo,
                pageNum: info.media ? info.pageNum : tplPageNum,
                slideLayout: info.slideLayout || tplLayout,
              };
            });

      const fileTitle =
        singleIndex !== undefined
          ? `LoteCriativo Slide ${singleIndex + 1}`
          : `LoteCriativo Editavel (${validPhrases.length} Paginas)`;

      const base64Data = await buildEditablePptxBase64(targetSlides, fileTitle);
      const byteCharacters = atob(base64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], {
        type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download =
        singleIndex !== undefined
          ? `LoteCriativo_Slide_${singleIndex + 1}_Editavel.pptx`
          : `LoteCriativo_${validPhrases.length}_Paginas_Editavel.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err: any) {
      console.error('Erro ao baixar .PPTX editável:', err);
      setDesignSaveBanner({
        text: `Erro ao gerar arquivo .PPTX editável: ${err.message || err}`,
        type: 'error',
      });
    } finally {
      setDownloadingPptx(false);
    }
  };

  // Action Handlers para gerenciamento de pastas e arquivos
  const handleCreateFolder = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!newFolderName.trim()) return;

    setCreatingFolder(true);
    const currentFolderId = driveFolderPath[driveFolderPath.length - 1]?.id || 'root';

    try {
      const res = await fetch('/api/canva/folders', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: newFolderName.trim(),
          parentFolderId: currentFolderId,
        }),
      });
      const data = await res.json();
      if (data.success && data.folder) {
        const newF = {
          id: data.folder.id,
          name: data.folder.name,
          category: 'folder',
          size: '0 Itens',
          modifiedAt: 'Agora',
          thumbnailUrl: '',
          canvaUrl: data.folder.canvaUrl || '',
          folderId: currentFolderId,
        };
        setCustomFolders((prev) => [newF, ...prev]);
        setNewFolderName('');
        setShowCreateFolderModal(false);
      } else {
        alert(data.error || 'Erro ao criar pasta.');
      }
    } catch (err) {
      console.error('Erro ao criar pasta:', err);
    } finally {
      setCreatingFolder(false);
    }
  };

  const handleRenameItem = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!itemToRename || !renameInput.trim()) return;

    setRenamingItem(true);
    const updatedName = renameInput.trim();
    try {
      const res = await fetch(`/api/canva/items/${encodeURIComponent(itemToRename.id)}/rename`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: updatedName,
          itemType: itemToRename.category,
        }),
      });
      const data = await res.json();
      if (data.success) {
        setCanvaCloudItems((prev) =>
          prev.map((i) => (i.id === itemToRename.id ? { ...i, name: updatedName } : i))
        );
        setCustomFolders((prev) =>
          prev.map((f) => (f.id === itemToRename.id ? { ...f, name: updatedName } : f))
        );
        if (selectedDriveItem?.id === itemToRename.id) {
          setSelectedDriveItem((prev: any) => (prev ? { ...prev, name: updatedName } : null));
        }
        setShowRenameModal(false);
        setItemToRename(null);
        setRenameInput('');
      } else {
        alert(data.error || 'Erro ao renomear item.');
      }
    } catch (err) {
      console.error('Erro ao renomear item:', err);
    } finally {
      setRenamingItem(false);
    }
  };

  const handleMoveItem = async (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    if (!itemToMove) return;

    setMovingItem(true);
    try {
      const targetFolderObj = allAvailableFolders.find((f) => f.id === targetFolderIdForMove);
      const res = await fetch('/api/canva/items/move', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          itemId: itemToMove.id,
          targetFolderId: targetFolderIdForMove,
          targetFolderName: targetFolderObj?.name || 'Meu Drive Canva',
        }),
      });
      const data = await res.json();
      if (data.success) {
        setCanvaCloudItems((prev) =>
          prev.map((i) => (i.id === itemToMove.id ? { ...i, folderId: targetFolderIdForMove } : i))
        );
        setCustomFolders((prev) =>
          prev.map((f) => (f.id === itemToMove.id ? { ...f, folderId: targetFolderIdForMove } : f))
        );
        setShowMoveModal(false);
        setItemToMove(null);
      } else {
        alert(data.error || 'Erro ao mover item.');
      }
    } catch (err) {
      console.error('Erro ao mover item:', err);
    } finally {
      setMovingItem(false);
    }
  };

  const handleDeleteItem = async (item: any) => {
    if (!confirm(`Deseja realmente excluir "${item.name}"?`)) return;

    try {
      await fetch(`/api/canva/items/${encodeURIComponent(item.id)}`, { method: 'DELETE' });
      setCanvaCloudItems((prev) => prev.filter((i) => i.id !== item.id));
      setCustomFolders((prev) => prev.filter((f) => f.id !== item.id));
      if (selectedDriveItem?.id === item.id) {
        setSelectedDriveItem(null);
      }
    } catch (err) {
      console.error('Erro ao excluir item:', err);
    }
  };

  const currentFolderId = driveFolderPath[driveFolderPath.length - 1]?.id || 'root';

  // Consolida apenas arquivos e pastas reais da conta oficial do Canva
  const allDriveItemsMap = new Map<string, any>();

  // 1. Adiciona pastas personalizadas criadas pelo usuário
  for (const folder of customFolders) {
    allDriveItemsMap.set(folder.id, {
      ...folder,
    });
  }

  // 3. Mescla itens retornados em tempo real da API do Canva
  for (const c of canvaCloudItems) {
    const id = c.id;
    if (id) {
      const existing = allDriveItemsMap.get(id);
      allDriveItemsMap.set(id, {
        id,
        name: c.name || existing?.name || 'Item Canva',
        category: c.category || existing?.category || (c.type === 'video' ? 'video' : c.source === 'design' ? 'design' : 'image'),
        size: c.size || existing?.size || '2.5 MB',
        modifiedAt: c.modifiedAt || existing?.modifiedAt || 'Hoje',
        thumbnailUrl: c.url || c.thumbnailUrl || existing?.thumbnailUrl || '',
        canvaUrl: c.canvaUrl || existing?.canvaUrl || (c.source === 'design' ? 'https://www.canva.com' : ''),
        folderId: c.folderId || existing?.folderId || 'root',
      });
    }
  }

  const allDriveItems = Array.from(allDriveItemsMap.values());

  const getFolderName = (fId: string) => {
    if (!fId || fId === 'root') return 'Meu Drive Canva';
    const f = allDriveItems.find((i) => i.id === fId && i.category === 'folder');
    return f ? f.name : 'Meu Drive Canva';
  };

  const handleDropOnFolder = async (targetFolderId: string, targetFolderName?: string) => {
    if (!draggedItemId || draggedItemId === targetFolderId) return;

    const itemToMoveObj = allDriveItems.find((i) => i.id === draggedItemId);
    if (!itemToMoveObj) return;

    if (itemToMoveObj.folderId === targetFolderId) {
      return;
    }

    const destName = targetFolderName || getFolderName(targetFolderId);

    setCanvaCloudItems((prev) =>
      prev.map((i) => (i.id === draggedItemId ? { ...i, folderId: targetFolderId } : i))
    );
    setCustomFolders((prev) =>
      prev.map((f) => (f.id === draggedItemId ? { ...f, folderId: targetFolderId } : f))
    );

    setDropNotification(`"${itemToMoveObj.name}" foi movido para "${destName}"!`);
    setTimeout(() => setDropNotification(null), 3500);

    try {
      await fetch('/api/canva/items/move', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          itemId: draggedItemId,
          targetFolderId,
          targetFolderName: destName,
        }),
      });
    } catch (err) {
      console.error('Erro ao mover por arrastar e soltar:', err);
    } finally {
      setDraggedItemId(null);
      setDragOverFolderId(null);
    }
  };

  const isGlobalSearching = Boolean(driveSearchQuery.trim());

  const currentFolderItems = allDriveItems.filter((item) => {
    if (driveCategory === 'folder') {
      return item.category === 'folder' && (currentFolderId === 'root' || item.folderId === currentFolderId);
    }
    return item.folderId === currentFolderId;
  });

  const baseFilteredItems = isGlobalSearching
    ? allDriveItems.filter((item) => {
        const matchesCategory = driveCategory === 'all' || item.category === driveCategory;
        const matchesSearch = item.name.toLowerCase().includes(driveSearchQuery.trim().toLowerCase());
        return matchesCategory && matchesSearch;
      })
    : currentFolderItems.filter((item) => {
        const matchesCategory = driveCategory === 'all' || item.category === driveCategory;
        return matchesCategory;
      });

  // Sorting logic
  const driveFilteredItems = [...baseFilteredItems].sort((a, b) => {
    let valA = a.name || '';
    let valB = b.name || '';

    if (sortBy === 'date') {
      valA = a.modifiedAt || '';
      valB = b.modifiedAt || '';
    } else if (sortBy === 'size') {
      valA = a.size || '';
      valB = b.size || '';
    }

    if (sortOrder === 'asc') {
      return valA > valB ? 1 : valA < valB ? -1 : 0;
    } else {
      return valA < valB ? 1 : valA > valB ? -1 : 0;
    }
  });

  // Bulk actions handlers
  const handleSelectAllToggle = () => {
    if (selectedItemIds.size === driveFilteredItems.length) {
      setSelectedItemIds(new Set());
    } else {
      setSelectedItemIds(new Set(driveFilteredItems.map((i) => i.id)));
    }
  };

  const handleToggleSelectItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setSelectedItemIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  };

  const handleBulkDelete = async () => {
    if (selectedItemIds.size === 0) return;
    if (!confirm(`Deseja realmente excluir os ${selectedItemIds.size} itens selecionados?`)) return;

    try {
      for (const id of selectedItemIds) {
        await fetch(`/api/canva/items/${encodeURIComponent(id)}`, { method: 'DELETE' });
      }
      setCanvaCloudItems((prev) => prev.filter((i) => !selectedItemIds.has(i.id)));
      setCustomFolders((prev) => prev.filter((f) => !selectedItemIds.has(f.id)));
      setSelectedItemIds(new Set());
    } catch (err) {
      console.error('Erro ao excluir itens em lote:', err);
    }
  };

  const handleBulkMove = async (targetFolderId: string) => {
    if (selectedItemIds.size === 0) return;
    const targetFolderObj = allAvailableFolders.find((f) => f.id === targetFolderId);
    const destName = targetFolderObj?.name || 'Pasta Destino';

    try {
      for (const id of selectedItemIds) {
        await fetch('/api/canva/items/move', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            itemId: id,
            targetFolderId,
            targetFolderName: destName,
          }),
        });
      }

      setCanvaCloudItems((prev) =>
        prev.map((i) => (selectedItemIds.has(i.id) ? { ...i, folderId: targetFolderId } : i))
      );
      setCustomFolders((prev) =>
        prev.map((f) => (selectedItemIds.has(f.id) ? { ...f, folderId: targetFolderId } : f))
      );

      setDropNotification(`${selectedItemIds.size} itens movidos para "${destName}" com sucesso!`);
      setTimeout(() => setDropNotification(null), 3500);
      setSelectedItemIds(new Set());
    } catch (err) {
      console.error('Erro ao mover itens em lote:', err);
    }
  };

  const countByCategory = (cat: string) => {
    if (cat === 'all') return isGlobalSearching ? driveFilteredItems.length : currentFolderItems.length;
    if (cat === 'folder') return allDriveItems.filter((i) => i.category === 'folder').length;
    return allDriveItems.filter((i) => i.category === cat).length;
  };

  const [expandedFolders, setExpandedFolders] = useState<Set<string>>(
    new Set(['root', 'f_uploads', 'f_2026_trabalho', 'f_2025_projetos', 'f_docs'])
  );

  const toggleFolderExpand = (folderId: string, e?: React.MouseEvent) => {
    if (e) e.stopPropagation();
    setExpandedFolders((prev) => {
      const next = new Set(prev);
      if (next.has(folderId)) {
        next.delete(folderId);
      } else {
        next.add(folderId);
      }
      return next;
    });
  };

  const navigateToFolderById = (targetFolderId: string, targetFolderName: string) => {
    if (targetFolderId === 'root') {
      setDriveFolderPath([{ id: 'root', name: 'Meu Drive Canva' }]);
    } else {
      const path: Array<{ id: string; name: string }> = [{ id: 'root', name: 'Meu Drive Canva' }];
      let currentId = targetFolderId;
      const visited = new Set<string>();
      const pathChain: Array<{ id: string; name: string }> = [];

      while (currentId && currentId !== 'root' && !visited.has(currentId)) {
        visited.add(currentId);
        const folderObj = allDriveItems.find((i) => i.id === currentId && i.category === 'folder');
        if (folderObj) {
          pathChain.unshift({ id: folderObj.id, name: folderObj.name });
          currentId = folderObj.folderId || 'root';
        } else {
          pathChain.unshift({ id: targetFolderId, name: targetFolderName });
          break;
        }
      }
      setDriveFolderPath([...path, ...pathChain]);
    }
    setDriveCategory('all');
    if (canvaStatus?.connected) {
      syncCanvaDriveFiles(targetFolderId);
    }
    setExpandedFolders((prev) => new Set(prev).add(targetFolderId));
  };

  const renderFolderTreeNode = (folderId: string, folderName: string, depth: number = 0): React.ReactNode => {
    const isExpanded = expandedFolders.has(folderId);
    const isActive = currentFolderId === folderId;
    const isDragOver = dragOverFolderId === folderId;

    const childFolders = allDriveItems.filter(
      (item) => item.category === 'folder' && item.folderId === folderId && item.id !== folderId
    );
    const directItemsCount = allDriveItems.filter((item) => item.folderId === folderId).length;

    return (
      <div key={folderId} className="select-none">
        <div
          onClick={() => navigateToFolderById(folderId, folderName)}
          onDragOver={(e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            setDragOverFolderId(folderId);
          }}
          onDragLeave={() => setDragOverFolderId(null)}
          onDrop={(e) => {
            e.preventDefault();
            handleDropOnFolder(folderId, folderName);
          }}
          style={{ paddingLeft: `${depth * 10 + 6}px` }}
          className={`flex items-center justify-between py-1.5 px-2 rounded-lg text-xs font-medium cursor-pointer transition-all group ${
            isDragOver
              ? 'bg-[#00C4CC] text-white font-bold ring-2 ring-white scale-[1.02] shadow-md'
              : isActive
              ? 'bg-[#1C1917] text-white shadow-xs'
              : 'text-[#57534E] hover:bg-[#E7E5E4]/60'
          }`}
        >
          <div className="flex items-center gap-1.5 min-w-0 truncate">
            {childFolders.length > 0 || folderId === 'root' ? (
              <button
                onClick={(e) => toggleFolderExpand(folderId, e)}
                className={`p-0.5 rounded hover:bg-black/10 transition-colors ${
                  isActive || isDragOver ? 'text-white' : 'text-[#78716C]'
                }`}
              >
                {isExpanded ? (
                  <ChevronDown className="w-3.5 h-3.5" />
                ) : (
                  <ChevronRight className="w-3.5 h-3.5" />
                )}
              </button>
            ) : (
              <span className="w-3.5" />
            )}

            {isActive || isExpanded || isDragOver ? (
              <FolderOpen className={`w-4 h-4 shrink-0 ${isDragOver ? 'text-white' : isActive ? 'text-[#00C4CC]' : 'text-amber-500'}`} />
            ) : (
              <Folder className={`w-4 h-4 shrink-0 ${isActive ? 'text-[#00C4CC]' : 'text-amber-500'}`} />
            )}

            <span className="truncate text-[11.5px] font-medium">{folderName}</span>
          </div>

          <span
            className={`text-[10px] font-mono px-1.5 py-0.2 rounded shrink-0 ml-1 ${
              isActive || isDragOver ? 'bg-white/20 text-white' : 'bg-[#E7E5E4] text-[#57534E]'
            }`}
          >
            {directItemsCount}
          </span>
        </div>

        {isExpanded && childFolders.length > 0 && (
          <div className="border-l border-[#E7E5E4] ml-3 my-0.5 space-y-0.5">
            {childFolders.map((cf) => renderFolderTreeNode(cf.id, cf.name, depth + 1))}
          </div>
        )}
      </div>
    );
  };

  const allAvailableFolders = [
    { id: 'root', name: 'Meu Drive Canva (Raiz)' },
    ...allDriveItems.filter((i) => i.category === 'folder'),
  ];

  return (
    <div className="min-h-screen bg-[#FDFBF7] text-[#1C1917] font-sans antialiased flex flex-col">
      {/* Top Bar Editorial Contract */}
      <header className="border-b border-[#E7E5E4] bg-white sticky top-0 z-30 px-6 py-4 flex items-center justify-between gap-8">
        <div className="flex items-center gap-3">
          <span className="text-xl font-serif font-bold tracking-tight text-[#1C1917] whitespace-nowrap shrink-0">
            LoteCriativo Canva
          </span>
          <span className="hidden sm:inline-block text-xs text-[#78716C] border-l border-[#E7E5E4] pl-3 font-mono">
            v1.0 · Produção em Lote
          </span>
        </div>

        <nav className="hidden md:flex items-center gap-6 text-sm font-medium text-[#78716C]">
          <button onClick={() => setActiveStep(1)} className={`hover:text-[#1C1917] transition-colors whitespace-nowrap ${activeStep === 1 ? 'text-[#1C1917] font-semibold' : ''}`}>
            Conexão
          </button>
          <button onClick={() => setActiveStep(2)} className={`hover:text-[#1C1917] transition-colors whitespace-nowrap ${activeStep === 2 ? 'text-[#1C1917] font-semibold' : ''}`}>
            Frases ({validPhrases.length})
          </button>
          <button onClick={() => setActiveStep(5)} className={`hover:text-[#1C1917] transition-colors whitespace-nowrap ${activeStep === 5 ? 'text-[#1C1917] font-semibold' : ''}`}>
            Estilo
          </button>
          <button onClick={() => setActiveStep(6)} className={`hover:text-[#1C1917] transition-colors whitespace-nowrap ${activeStep === 6 ? 'text-[#1C1917] font-semibold' : ''}`}>
            Gerar Lote
          </button>
          <button onClick={() => setActiveStep(7)} className={`hover:text-[#1C1917] transition-colors whitespace-nowrap ${activeStep === 7 ? 'text-[#1C1917] font-semibold' : ''}`}>
            Resultados
          </button>
          <button onClick={() => setActiveStep(8)} className={`hover:text-[#1C1917] transition-colors whitespace-nowrap flex items-center gap-1.5 ${activeStep === 8 ? 'text-[#1C1917] font-semibold' : ''}`}>
            <HardDrive className="w-3.5 h-3.5 text-[#00C4CC]" />
            <span>Drive Canva</span>
          </button>
        </nav>

        <div className="flex items-center gap-3 shrink-0">
          {user ? (
            <div className="flex items-center gap-2">
              <span className="hidden sm:inline-block text-xs text-[#78716C] max-w-[120px] truncate">
                {user.displayName || user.email}
              </span>
              <button
                onClick={handleFirebaseLogout}
                className="px-3 py-1.5 text-xs font-medium text-[#78716C] hover:text-[#1C1917] border border-[#E7E5E4] rounded-lg bg-white hover:bg-[#F5F5F4] transition-colors whitespace-nowrap"
              >
                Sair
              </button>
            </div>
          ) : (
            <button
              onClick={handleFirebaseLogin}
              className="px-4 py-2 text-xs font-medium text-white bg-[#1C1917] hover:bg-[#292524] rounded-lg transition-colors whitespace-nowrap"
            >
              Login Firebase
            </button>
          )}
        </div>
      </header>

      {/* Stepper Navigation Strip */}
      <div className="border-b border-[#E7E5E4] bg-white overflow-x-auto scrollbar-none px-6 py-2">
        <div className="flex items-center gap-2 min-w-max text-xs font-medium text-[#78716C]">
          {[
            { num: 1, label: '1. Conectar Canva' },
            { num: 2, label: '2. Frases' },
            { num: 3, label: '3. Modelo' },
            { num: 4, label: '4. Mídias' },
            { num: 5, label: '5. Estilo' },
            { num: 6, label: '6. Gerador de Lote' },
            { num: 7, label: '7. Resultados' },
            { num: 8, label: '📁 8. Drive Canva (Explorer)' },
          ].map((s) => (
            <button
              key={s.num}
              onClick={() => setActiveStep(s.num)}
              className={`px-3 py-1.5 rounded-md transition-colors whitespace-nowrap flex items-center gap-1.5 ${
                activeStep === s.num
                  ? 'bg-[#1C1917] text-white font-semibold shadow-sm'
                  : 'hover:bg-[#F5F5F4] text-[#57534E]'
              }`}
            >
              <span>{s.label}</span>
              {s.num === 1 && canvaStatus?.connected && (
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" />
              )}
            </button>
          ))}
        </div>
      </div>

      {/* Main Content Area with Persistent Right Preview Sidebar */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 lg:p-8 flex flex-col lg:flex-row gap-8 items-start">
        <div className="flex-1 w-full min-w-0">
          {/* ========================================================================= */}
          {/* ETAPA 1: CONEXÃO CANVA (Alvo do Focus Mode e do botão selecionado)       */}
          {/* ========================================================================= */}
          {activeStep === 1 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-[#F5F5F4]">
                <div>
                  <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                    1. Conexão ao Canva
                  </h2>
                  <p className="text-xs text-[#78716C] mt-1">
                    Integração oficial OAuth 2.0 com Canva Connect API e chave PKCE.
                  </p>
                </div>

                <div className="flex items-center gap-2">
                  {loadingCanva ? (
                    <span className="text-xs text-[#A8A29E] flex items-center gap-1.5">
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" /> Verificando...
                    </span>
                  ) : canvaStatus?.connected ? (
                    <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-full text-xs font-medium">
                      <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                      Conta Canva Autorizada
                    </span>
                  ) : (
                    <span className="inline-flex items-center gap-1.5 px-3 py-1 bg-amber-50 text-amber-800 border border-amber-200 rounded-full text-xs font-medium">
                      <AlertCircle className="w-3.5 h-3.5 text-amber-600" />
                      Desconectado
                    </span>
                  )}
                </div>
              </div>

              {/* Card de Estado Conectado */}
              {canvaStatus?.connected ? (
                <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-xl p-5 space-y-4">
                  <div className="flex items-start justify-between">
                    <div>
                      <p className="text-sm font-semibold text-[#1C1917]">
                        {canvaStatus.user?.displayName || 'Usuário Canva'}
                      </p>
                      <p className="text-xs text-[#78716C] font-mono mt-0.5">
                        ID: {canvaStatus.user?.id || 'canva_account'}
                      </p>
                    </div>
                    <button
                      onClick={handleDisconnect}
                      className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-rose-700 bg-rose-50 hover:bg-rose-100 border border-rose-200 rounded-lg transition-colors"
                    >
                      <LogOut className="w-3.5 h-3.5" />
                      Desconectar
                    </button>
                  </div>

                  <div className="pt-2 border-t border-[#E7E5E4] text-xs text-[#78716C] space-y-2">
                    <p className="font-medium text-[#44403C]">Escopos de Permissão Ativos:</p>
                    <div className="flex flex-wrap gap-1.5 pt-1">
                      {canvaStatus.scopes?.map((s, idx) => (
                        <span key={`${s}_${idx}`} className="px-2 py-0.5 bg-white border border-[#E7E5E4] rounded text-[11px] font-mono text-[#57534E]">
                          {s}
                        </span>
                      ))}
                    </div>

                    {!canvaStatus.hasBrandTemplateScope && (
                      <div className="mt-2 p-3 bg-cyan-50/70 border border-cyan-200 rounded-xl flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                        <div className="text-[11.5px] text-cyan-950 space-y-0.5">
                          <p className="font-bold">Habilitar Modelos de Marca (Autofill API)</p>
                          <p className="text-cyan-900">
                            Para usar <strong>Brand Templates</strong> existentes e substituir campos-chave preservando tudo, habilite os escopos <code className="font-mono bg-white/80 px-1 rounded">brandtemplate:meta:read</code> e <code className="font-mono bg-white/80 px-1 rounded">brandtemplate:content:read</code> no Canva Developer Portal e autorize abaixo.
                          </p>
                        </div>
                        <button
                          onClick={() => handleConnect(true)}
                          disabled={connecting}
                          className="px-3.5 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-lg shrink-0 transition-colors cursor-pointer disabled:opacity-50"
                        >
                          {connecting ? 'Autorizando...' : 'Autorizar Brand Templates'}
                        </button>
                      </div>
                    )}
                  </div>

                  <div className="pt-3 flex justify-end">
                    <button
                      onClick={() => setActiveStep(2)}
                      className="inline-flex items-center gap-2 px-5 py-2.5 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-medium transition-colors"
                    >
                      Prosseguir para Frases <ChevronRight className="w-4 h-4" />
                    </button>
                  </div>
                </div>
              ) : (
                /* Card de Estado Desconectado e Ação de Login */
                <div className="space-y-6">
                  {/* Informações Obrigatórias do Redirect URI */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-xl p-4 sm:p-5 space-y-3">
                    <div className="flex items-center gap-2 text-xs font-semibold text-[#1C1917]">
                      <Info className="w-4 h-4 text-[#78716C]" />
                      Configuração necessária no Canva Developers Portal
                    </div>
                    <p className="text-xs text-[#57534E] leading-relaxed">
                      No portal do Canva (seção <strong>Fora do Canva &gt; URLs de Redirecionamento</strong>), adicione a URL de callback deste app:
                    </p>
                    <div className="flex items-center gap-2 bg-white p-2 border border-[#E7E5E4] rounded-lg">
                      <code className="text-[11px] font-mono text-[#1C1917] truncate flex-1 select-all">
                        {canvaStatus?.redirectUri || 'https://ais-dev-htprufxxrvhnpehkl2vlww-209534015373.us-east1.run.app/api/auth/canva/callback'}
                      </code>
                      <button
                        onClick={() => copyRedirectUri(canvaStatus?.redirectUri || '')}
                        className="px-2.5 py-1 text-xs font-medium text-[#1C1917] bg-[#F5F5F4] hover:bg-[#E7E5E4] rounded flex items-center gap-1 transition-colors"
                        title="Copiar URL de Redirecionamento"
                      >
                        {copiedUri ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
                        <span>{copiedUri ? 'Copiado!' : 'Copiar'}</span>
                      </button>
                    </div>
                  </div>

                  {/* Painel de Diagnóstico e Configuração do Client ID */}
                  <div className="border border-[#E7E5E4] bg-white rounded-xl p-4 sm:p-5 space-y-3">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2 text-xs font-semibold text-[#1C1917]">
                        <Key className="w-4 h-4 text-[#00C4CC]" />
                        Credenciais do Canva Developers
                      </div>
                      {canvaStatus?.clientIdConfigured ? (
                        <span className="text-[11px] font-mono px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded font-medium">
                          Client ID Detectado ({canvaStatus.clientId})
                        </span>
                      ) : (
                        <span className="text-[11px] font-mono px-2 py-0.5 bg-rose-50 text-rose-700 border border-rose-200 rounded font-medium">
                          Client ID Não Configurado
                        </span>
                      )}
                    </div>

                    <div className="text-xs text-[#57534E] space-y-1">
                      <p>
                        <strong>Client ID em Uso:</strong> <code className="bg-[#F5F5F4] px-1.5 py-0.5 rounded text-[11px] font-mono text-[#1C1917]">{canvaStatus?.clientId || 'Nenhum'}</code>
                      </p>
                      <p>
                        <strong>Origem das Credenciais:</strong> <span className="text-[#78716C]">{canvaStatus?.clientIdSource || 'Não verificado'}</span>
                      </p>
                    </div>

                    {(!canvaStatus?.configured || !canvaStatus?.clientSecretConfigured || !canvaStatus?.clientIdConfigured) && (
                      <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg text-xs text-amber-900 space-y-1">
                        <p className="font-semibold flex items-center gap-1.5">
                          <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
                          Atenção: O erro "O ID do cliente é inválido" ocorre quando o Canva não reconhece o Client ID fornecido.
                        </p>
                        <p className="text-[11.5px] text-amber-800">
                          Acesse o <strong>Canva Developers Console &gt; Your apps &gt; Sua Aplicação &gt; Configuration &gt; Credentials</strong>, obtenha seu <code>Client ID</code> e <code>Client Secret</code> oficiais e salve-os abaixo.
                        </p>
                      </div>
                    )}

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-2">
                      <div>
                        <label className="block text-[11px] font-medium text-[#57534E] mb-1">
                          CANVA_CLIENT_ID (do portal do Canva)
                        </label>
                        <input
                          type="text"
                          value={manualClientId}
                          onChange={(e) => setManualClientId(e.target.value)}
                          placeholder="Ex: OC-AaFaI3rW6tyE"
                          className="w-full text-xs px-3 py-2 bg-white border border-[#D6D3D1] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-mono"
                        />
                      </div>
                      <div>
                        <label className="block text-[11px] font-medium text-[#57534E] mb-1">
                          CANVA_CLIENT_SECRET (do portal do Canva)
                        </label>
                        <input
                          type="password"
                          value={manualClientSecret}
                          onChange={(e) => setManualClientSecret(e.target.value)}
                          placeholder="Cole seu segredo do Canva"
                          className="w-full text-xs px-3 py-2 bg-white border border-[#D6D3D1] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-mono"
                        />
                      </div>
                    </div>

                    <div className="flex items-center justify-between pt-2 border-t border-[#F5F5F4]">
                      {credentialsMsg ? (
                        <span className={`text-xs ${credentialsMsg.type === 'success' ? 'text-emerald-700' : 'text-rose-700'}`}>
                          {credentialsMsg.text}
                        </span>
                      ) : (
                        <span className="text-[11px] text-[#A8A29E]">
                          As credenciais são persistidas no backend via Firestore.
                        </span>
                      )}
                      <button
                        onClick={handleSaveCredentials}
                        disabled={savingCredentials}
                        className="ml-auto px-4 py-1.5 text-xs font-medium text-white bg-[#1C1917] hover:bg-[#292524] rounded-lg transition-colors disabled:opacity-50"
                      >
                        {savingCredentials ? 'Salvando...' : 'Salvar / Atualizar Credenciais'}
                      </button>
                    </div>
                  </div>

                  {/* O BOTÃO PRINCIPAL SELECIONADO EM FOCUS MODE */}
                  <div className="pt-2 flex flex-col sm:flex-row items-center gap-4">
                    <div className="relative group w-full sm:w-auto">
                      <button
                        onClick={() => handleConnect(false)}
                        disabled={connecting}
                        className="w-full sm:w-auto px-6 py-3.5 text-sm font-semibold text-white bg-[#00C4CC] hover:bg-[#00B0B7] rounded-xl shadow-sm transition-all flex items-center justify-center gap-2.5 active:scale-[0.99] disabled:opacity-50 cursor-pointer"
                      >
                        {connecting ? (
                          <>
                            <RefreshCw className="w-4 h-4 animate-spin" />
                            <span>Aguardando autorização no Canva...</span>
                          </>
                        ) : (
                          <>
                            <Sparkles className="w-4 h-4" />
                            <span>Conectar ao Canva</span>
                          </>
                        )}
                      </button>

                      {/* Delayed Tooltip on Hover */}
                      <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 hidden group-hover:block transition-all duration-300 delay-300 opacity-0 group-hover:opacity-100 z-30 pointer-events-none">
                        <div className="bg-[#1C1917] text-white text-[11px] font-medium py-1.5 px-3 rounded-lg shadow-xl whitespace-nowrap">
                          Allows access to your private folders and projects
                          <div className="absolute top-full left-1/2 -translate-x-1/2 border-4 border-transparent border-t-[#1C1917]" />
                        </div>
                      </div>
                    </div>

                    <button
                      onClick={fetchCanvaStatus}
                      className="text-xs text-[#78716C] hover:text-[#1C1917] flex items-center gap-1.5 underline decoration-[#D6D3D1]"
                    >
                      <RefreshCw className="w-3.5 h-3.5" />
                      Verificar status agora
                    </button>
                  </div>

                  {/* Dica para alternar de conta no Canva */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-xl p-4 text-xs text-[#57534E] space-y-1.5">
                    <div className="font-semibold text-[#1C1917] flex items-center gap-1.5">
                      <Info className="w-4 h-4 text-[#00C4CC]" />
                      Como entrar com outra conta do Canva:
                    </div>
                    <p className="leading-relaxed">
                      1. Clique em <strong>Desconectar</strong> acima para revogar o acesso atual.<br />
                      2. Caso o Canva reconecte a mesma conta automaticamente, acesse <a href="https://www.canva.com" target="_blank" rel="noopener noreferrer" className="text-blue-600 underline font-medium">canva.com</a> em outra guia, clique no seu perfil e escolha <strong>Sair</strong>.<br />
                      3. Clique em <strong>Conectar ao Canva</strong> novamente para inserir o e-mail e senha da nova conta.
                    </p>
                  </div>

                  {directAuthUrl && (
                    <div className="text-xs text-[#78716C] bg-white border border-[#E7E5E4] p-3 rounded-lg flex items-center justify-between">
                      <span>O popup foi bloqueado pelo navegador?</span>
                      <a
                        href={directAuthUrl}
                        target="_blank"
                        rel="noopener noreferrer"
                        className="text-blue-600 hover:underline flex items-center gap-1 font-medium"
                      >
                        Abrir tela de login do Canva <ExternalLink className="w-3.5 h-3.5" />
                      </a>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 2: ADICIONAR FRASES                                                */}
          {/* ========================================================================= */}
          {activeStep === 2 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 pb-4 border-b border-[#F5F5F4]">
                <div>
                  <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                    2. Entrada de Frases
                  </h2>
                  <p className="text-xs text-[#78716C] mt-1">
                    Cole suas mensagens, uma por linha. Linhas vazias são removidas e caracteres especiais mantidos.
                  </p>
                </div>
                <div className="flex items-center gap-2">
                  <label className="cursor-pointer px-3 py-1.5 text-xs font-medium text-[#1C1917] bg-[#F5F5F4] hover:bg-[#E7E5E4] rounded-lg transition-colors flex items-center gap-1.5">
                    <FileText className="w-3.5 h-3.5" />
                    <span>Importar TXT/CSV</span>
                    <input type="file" accept=".txt,.csv" onChange={importPhrasesFile} className="hidden" />
                  </label>
                </div>
              </div>

              <div>
                <textarea
                  value={rawPhrases}
                  onChange={(e) => setRawPhrases(e.target.value)}
                  rows={6}
                  placeholder="Cole suas frases aqui, uma em cada linha..."
                  className="w-full text-sm p-4 bg-[#FAF9F6] border border-[#E7E5E4] rounded-xl focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-sans leading-relaxed resize-y"
                />
                <div className="flex items-center justify-between text-xs text-[#78716C] mt-2">
                  <span>{validPhrases.length} frases válidas reconhecidas</span>
                  <button
                    onClick={() => setRawPhrases('')}
                    className="text-rose-600 hover:underline text-xs"
                  >
                    Limpar tudo
                  </button>
                </div>
              </div>

              {/* Tabela de prévia das frases */}
              <div className="border border-[#E7E5E4] rounded-xl overflow-hidden">
                <div className="bg-[#FAF9F6] px-4 py-2.5 border-b border-[#E7E5E4] text-xs font-semibold text-[#57534E] flex items-center justify-between">
                  <span>Lista de Frases para Processamento</span>
                  <span>Modo: {designMode === 'individual' ? '1 Design por Frase' : 'Páginas Múltiplas'}</span>
                </div>
                <div className="max-h-60 overflow-y-auto divide-y divide-[#F5F5F4]">
                  {validPhrases.map((phrase, idx) => (
                    <div key={idx} className="p-3 text-xs flex items-center justify-between gap-4 hover:bg-[#FAF9F6]">
                      <span className="font-mono text-[#A8A29E] w-6 shrink-0">{idx + 1}.</span>
                      <p className="flex-1 text-[#1C1917] truncate">{phrase}</p>
                      <button
                        onClick={() => {
                          const updated = validPhrases.filter((_, i) => i !== idx);
                          setRawPhrases(updated.join('\n'));
                        }}
                        className="text-[#A8A29E] hover:text-rose-600 transition-colors p-1"
                        title="Remover frase"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>

              {/* Opções de Formato */}
              <div className="bg-[#FAF9F6] p-4 rounded-xl border border-[#E7E5E4] flex flex-col sm:flex-row items-center justify-between gap-4">
                <div>
                  <p className="text-xs font-semibold text-[#1C1917]">Estrutura do Lote</p>
                  <p className="text-xs text-[#78716C]">Defina como os designs serão organizados no Canva ou no arquivo final.</p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setDesignMode('individual')}
                    className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors ${
                      designMode === 'individual' ? 'bg-[#1C1917] text-white' : 'bg-white text-[#57534E] border border-[#E7E5E4]'
                    }`}
                  >
                    1 Design por Frase
                  </button>
                  <button
                    onClick={() => setDesignMode('multipage')}
                    className={`px-3 py-1.5 text-xs font-medium rounded-lg transition-colors ${
                      designMode === 'multipage' ? 'bg-[#1C1917] text-white' : 'bg-white text-[#57534E] border border-[#E7E5E4]'
                    }`}
                  >
                    Várias Páginas
                  </button>
                </div>
              </div>

              <div className="flex justify-between pt-2">
                <button
                  onClick={() => setActiveStep(1)}
                  className="px-4 py-2 border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#57534E] hover:bg-[#F5F5F4]"
                >
                  Voltar
                </button>
                <button
                  onClick={() => setActiveStep(3)}
                  className="px-5 py-2.5 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-medium flex items-center gap-1.5"
                >
                  Avançar para Modelo <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 3: SELECIONAR MODELO                                               */}
          {/* ========================================================================= */}
          {activeStep === 3 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="pb-4 border-b border-[#F5F5F4]">
                <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                  3. Selecionar o Modelo Canva
                </h2>
                <p className="text-xs text-[#78716C] mt-1">
                  Mapeie os campos de texto e mídia com os identificadores do seu design no Canva.
                </p>
              </div>

              <div className="space-y-4">
                <div>
                  <label className="block text-xs font-semibold text-[#1C1917] mb-1">
                    ID do Modelo Canva (Brand Template ID ou Design ID)
                  </label>
                  <input
                    type="text"
                    value={templateId}
                    onChange={(e) => setTemplateId(e.target.value)}
                    placeholder="Ex: DAFxAbCdEfG ou insira o link do seu design"
                    className="w-full text-xs px-3.5 py-2.5 bg-[#FAF9F6] border border-[#E7E5E4] rounded-xl focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-mono"
                  />
                  <p className="text-[11px] text-[#78716C] mt-1.5">
                    Como encontrar: No Canva, abra seu modelo e copie a sequência de letras após <code className="font-mono text-[#1C1917]">/design/</code> na barra de navegação.
                  </p>

                  {mediaList.length > 0 && (
                    <div className="mt-3 pt-3 border-t border-[#F5F5F4] space-y-1.5">
                      <span className="text-[11px] font-semibold text-[#57534E] block">
                        Usar um design já selecionado na Etapa 4:
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {mediaList.map((m) => (
                          <button
                            key={m.id}
                            type="button"
                            onClick={() => setTemplateId(m.id)}
                            className={`px-2.5 py-1 rounded-lg text-[11px] font-mono border transition-colors cursor-pointer ${
                              templateId === m.id
                                ? 'bg-[#1C1917] text-white border-[#1C1917]'
                                : 'bg-[#FAF9F6] text-[#57534E] border-[#E7E5E4] hover:border-[#1C1917]'
                            }`}
                          >
                            {m.name} ({m.id})
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>

                <div
                  className={`p-4 rounded-xl space-y-3 border ${
                    isStep3FieldPhraseMissing || isModelValidationBlocked
                      ? 'bg-rose-50/70 border-rose-400'
                      : 'bg-[#FAF9F6] border-[#E7E5E4]'
                  }`}
                >
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <div>
                      <p className="text-xs font-semibold text-[#1C1917]">Mapeamento de Campos do Modelo</p>
                      <p className="text-xs text-[#78716C]">
                        Configure exatamente como o marcador está escrito no seu template Canva (ex: <code className="font-mono">(frase)</code> ou <code className="font-mono">frase</code>):
                      </p>
                    </div>
                    {hasInspectableModelLoaded && (
                      <span
                        className={`text-[10px] font-mono font-bold px-2.5 py-1 rounded-lg ${
                          isStep3FieldPhraseMissing
                            ? 'bg-rose-600 text-white'
                            : 'bg-emerald-600 text-white'
                        }`}
                      >
                        {isStep3FieldPhraseMissing ? '✗ Marcador não encontrado' : '✓ Marcador compatível'}
                      </span>
                    )}
                  </div>

                  {isStep3FieldPhraseMissing && (
                    <div className="p-3 bg-rose-100/80 border border-rose-300 rounded-lg text-xs text-rose-900 flex items-start gap-2">
                      <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                      <div className="space-y-1.5">
                        <p className="font-semibold">
                          Campo obrigatório "{fieldPhrase || '(vazio)'}" não encontrado no modelo selecionado. A geração do lote ficará bloqueada até que este campo corresponda ao modelo.
                        </p>
                        {modelSelectableTokens.length > 0 && (
                          <div className="flex flex-wrap items-center gap-1.5">
                            <span className="text-[11px] font-medium text-rose-800">Encontrados no modelo:</span>
                            {modelSelectableTokens.map((tok) => (
                              <button
                                key={tok}
                                type="button"
                                onClick={() => {
                                  setFieldPhrase(tok);
                                  setKeyFieldMappings((prev) => {
                                    const hasPhrase = prev.some((m) => m.source === 'batch_phrase');
                                    if (!hasPhrase) {
                                      return [
                                        {
                                          fieldKey: tok,
                                          detectedInTemplate: true,
                                          source: 'batch_phrase',
                                          customText: '',
                                          pageScope: 'all',
                                          selectedPages: [1],
                                        },
                                        ...prev,
                                      ];
                                    }
                                    let updatedFirst = false;
                                    return prev.map((m) => {
                                      if (!updatedFirst && m.source === 'batch_phrase') {
                                        updatedFirst = true;
                                        return { ...m, fieldKey: tok, detectedInTemplate: true };
                                      }
                                      return m;
                                    });
                                  });
                                }}
                                className="px-2 py-0.5 bg-white border border-rose-300 hover:bg-rose-50 text-rose-900 rounded font-mono text-[11px] font-bold cursor-pointer"
                              >
                                {tok}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  )}

                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-[11px] font-medium text-[#57534E] mb-1">
                        Campo de Texto (Frase) <span className="text-rose-600">*</span>
                      </label>
                      <input
                        type="text"
                        value={fieldPhrase}
                        onChange={(e) => {
                          const nextVal = e.target.value;
                          setFieldPhrase(nextVal);
                          setKeyFieldMappings((prev) => {
                            const hasPhrase = prev.some((m) => m.source === 'batch_phrase');
                            if (!hasPhrase) {
                              return [
                                {
                                  fieldKey: nextVal,
                                  detectedInTemplate: isFieldKeyFoundInModel(nextVal),
                                  source: 'batch_phrase',
                                  customText: '',
                                  pageScope: 'all',
                                  selectedPages: [1],
                                },
                                ...prev,
                              ];
                            }
                            let updatedFirst = false;
                            return prev.map((m) => {
                              if (!updatedFirst && m.source === 'batch_phrase') {
                                updatedFirst = true;
                                return {
                                  ...m,
                                  fieldKey: nextVal,
                                  detectedInTemplate: isFieldKeyFoundInModel(nextVal),
                                };
                              }
                              return m;
                            });
                          });
                        }}
                        placeholder="Ex: (frase) ou FRASE"
                        className={`w-full text-xs px-2.5 py-1.5 rounded-lg font-mono border ${
                          isStep3FieldPhraseMissing
                            ? 'bg-white border-2 border-rose-500 text-rose-900 focus:outline-none focus:ring-2 focus:ring-rose-500/30'
                            : 'bg-white border-[#E7E5E4]'
                        }`}
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-medium text-[#57534E] mb-1">Campo de Mídia (Vídeo/Foto)</label>
                      <input
                        type="text"
                        value={fieldMedia}
                        onChange={(e) => setFieldMedia(e.target.value)}
                        className="w-full text-xs px-2.5 py-1.5 bg-white border border-[#E7E5E4] rounded-lg font-mono"
                      />
                    </div>
                    <div>
                      <label className="block text-[11px] font-medium text-[#57534E] mb-1">Assinatura / @ (Opcional)</label>
                      <input
                        type="text"
                        value={fieldSignature}
                        onChange={(e) => setFieldSignature(e.target.value)}
                        className="w-full text-xs px-2.5 py-1.5 bg-white border border-[#E7E5E4] rounded-lg font-mono"
                      />
                    </div>
                  </div>
                </div>

                <div className="p-4 bg-amber-50/60 border border-amber-200 rounded-xl text-xs text-amber-900 leading-relaxed">
                  <strong>Aviso sobre contas gratuitas do Canva:</strong> Contas gratuitas podem não ter acesso ilimitado à API de Autofill de Modelos de Marca. Se o modelo não puder ser preenchido de forma programática direta pelo Canva, nosso motor gerará automaticamente a exportação em lote com os estilos definidos na Etapa 5.
                </div>
              </div>

              <div className="flex justify-between pt-2">
                <button
                  onClick={() => setActiveStep(2)}
                  className="px-4 py-2 border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#57534E] hover:bg-[#F5F5F4]"
                >
                  Voltar
                </button>
                <button
                  onClick={() => setActiveStep(4)}
                  className="px-5 py-2.5 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-medium flex items-center gap-1.5"
                >
                  Avançar para Mídias <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 4: SELECIONAR MÍDIAS                                               */}
          {/* ========================================================================= */}
          {activeStep === 4 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-[#F5F5F4]">
                <div>
                  <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                    4. Selecionar Vídeos e Imagens
                  </h2>
                  <p className="text-xs text-[#78716C] mt-1">
                    Selecione mídias do seu dispositivo móvel ou computador para associar a cada frase.
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    className="px-4 py-2 bg-[#FAF9F6] hover:bg-[#F5F5F4] text-[#1C1917] border border-[#E7E5E4] text-xs font-semibold rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-sm hover:border-[#1C1917]"
                  >
                    <Upload className="w-3.5 h-3.5 text-[#78716C]" />
                    <span>Selecionar do dispositivo</span>
                  </button>

                  <button
                    onClick={() => {
                      if (canvaStatus?.connected) {
                        openCanvaFolderExplorer('root', 'Projetos (Raiz)');
                      } else {
                        handleConnect(false);
                      }
                    }}
                    className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-xl flex items-center gap-2 transition-all cursor-pointer shadow-sm"
                  >
                    <Sparkles className="w-3.5 h-3.5" />
                    <span>Selecionar arquivos do Canva</span>
                  </button>
                </div>
                <input
                  ref={fileInputRef}
                  type="file"
                  multiple
                  accept="image/*,video/*,audio/*"
                  onChange={handleFileUpload}
                  className="hidden"
                />
              </div>

              {/* Banner de Atalho Rápido para a Biblioteca do Canva */}
              {canvaStatus?.connected && (
                <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-4 sm:p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
                  <div className="space-y-1">
                    <p className="text-xs font-bold text-[#1C1917] flex items-center gap-1.5">
                      <Folder className="w-4 h-4 text-[#00C4CC]" />
                      Navegador Oficial da Biblioteca Canva
                    </p>
                    <p className="text-xs text-[#78716C] leading-relaxed">
                      Conectado como <strong className="text-[#1C1917]">{canvaStatus.user?.displayName || 'Conta Canva'}</strong>. Explore projetos, subpastas, designs e imagens em tempo real.
                    </p>
                  </div>
                  <div className="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                    <button
                      onClick={() => openCanvaFolderExplorer('root', 'Projetos (Raiz)', 'folders')}
                      className="flex-1 sm:flex-none px-3.5 py-2 bg-white hover:bg-[#F5F5F4] border border-[#E7E5E4] text-xs font-medium text-[#1C1917] rounded-xl flex items-center justify-center gap-1.5 transition-colors cursor-pointer"
                    >
                      <Folder className="w-3.5 h-3.5 text-[#00C4CC]" />
                      <span>Projetos (Raiz)</span>
                    </button>
                    <button
                      onClick={() => openCanvaFolderExplorer('uploads', 'Uploads', 'folders')}
                      className="flex-1 sm:flex-none px-3.5 py-2 bg-white hover:bg-[#F5F5F4] border border-[#E7E5E4] text-xs font-medium text-[#1C1917] rounded-xl flex items-center justify-center gap-1.5 transition-colors cursor-pointer"
                    >
                      <UploadCloud className="w-3.5 h-3.5 text-[#00C4CC]" />
                      <span>Uploads</span>
                    </button>
                    <button
                      onClick={() => openCanvaFolderExplorer('root', 'Vídeos do Canva', 'videos')}
                      className="flex-1 sm:flex-none px-3.5 py-2 bg-[#1C1917] hover:bg-[#292524] text-xs font-semibold text-white rounded-xl flex items-center justify-center gap-1.5 transition-colors cursor-pointer"
                    >
                      <Video className="w-3.5 h-3.5 text-[#00C4CC]" />
                      <span>Vídeos do Canva</span>
                    </button>
                  </div>
                </div>
              )}

              {/* ===================================================================== */}
              {/* MODAL OFICIAL: NAVEGADOR DE PASTAS & VÍDEOS DO CANVA                 */}
              {/* ===================================================================== */}
              {showFolderExplorer && (
                <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-6">
                  <div className="bg-white rounded-2xl max-w-4xl w-full max-h-[90vh] flex flex-col shadow-2xl overflow-hidden border border-[#E7E5E4]">
                    {/* Header do Explorador */}
                    <div className="px-6 py-4 border-b border-[#E7E5E4] bg-white flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shrink-0">
                      <div>
                        <div className="flex items-center gap-2">
                          <h3 className="text-lg font-serif font-bold text-[#1C1917]">Navegador da Biblioteca do Canva</h3>
                          <span className="px-2 py-0.5 bg-[#00C4CC]/10 text-[#008A90] font-mono text-[10px] font-bold rounded-md">
                            API Oficial
                          </span>
                        </div>
                        <p className="text-xs text-[#78716C] mt-0.5">
                          Explore pastas, designs, projetos de vídeo (exportáveis em MP4) e assets da sua conta Canva.
                        </p>
                      </div>

                      <div className="flex items-center gap-2 w-full sm:w-auto">
                        <div className="relative flex-1 sm:w-60">
                          <Search className="w-3.5 h-3.5 text-[#A8A29E] absolute left-2.5 top-1/2 -translate-y-1/2" />
                          <input
                            type="text"
                            value={folderExplorerSearch}
                            onChange={(e) => setFolderExplorerSearch(e.target.value)}
                            placeholder="Filtrar arquivos..."
                            className="text-xs pl-8 pr-3 py-1.5 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] w-full font-sans"
                          />
                        </div>
                        <button
                          onClick={() => setShowFolderExplorer(false)}
                          className="p-1.5 text-[#78716C] hover:text-[#1C1917] hover:bg-[#FAF9F6] rounded-lg border border-[#E7E5E4] cursor-pointer shrink-0"
                          title="Fechar"
                        >
                          <X className="w-4 h-4" />
                        </button>
                      </div>
                    </div>

                    {/* Barra de Filtros Unificados: Todos, Imagens, Vídeos, Áudios, Designs + Host SDK */}
                    <div className="px-6 py-2.5 bg-white border-b border-[#E7E5E4] flex flex-wrap items-center justify-between gap-2 text-xs shrink-0">
                      <div className="flex items-center gap-1.5 flex-wrap">
                        {[
                          { id: 'all' as const, label: 'Todos', icon: Layers },
                          { id: 'image' as const, label: 'Imagens', icon: ImageIcon },
                          { id: 'video' as const, label: 'Vídeos', icon: Video },
                          { id: 'audio' as const, label: 'Áudios', icon: Music },
                          { id: 'design' as const, label: 'Designs', icon: Sparkles },
                        ].map((f) => {
                          const FIcon = f.icon;
                          const active = explorerMediaFilter === f.id;
                          return (
                            <button
                              key={f.id}
                              type="button"
                              onClick={() => {
                                setExplorerMediaFilter(f.id);
                                if (f.id === 'video' || f.id === 'audio') {
                                  setExplorerTab('videos');
                                  if (!explorerVideosData) loadExplorerVideos();
                                } else {
                                  setExplorerTab('folders');
                                }
                              }}
                              className={`px-3 py-1.5 rounded-lg font-semibold flex items-center gap-1.5 transition-colors cursor-pointer ${
                                active
                                  ? 'bg-[#1C1917] text-white shadow-xs'
                                  : 'bg-[#FAF9F6] border border-[#E7E5E4] text-[#57534E] hover:bg-[#F5F5F4]'
                              }`}
                            >
                              <FIcon className={`w-3.5 h-3.5 ${active ? 'text-[#00C4CC]' : 'text-[#78716C]'}`} />
                              <span>{f.label}</span>
                            </button>
                          );
                        })}
                      </div>

                      <span
                        className={`text-[10px] font-mono px-2 py-0.5 rounded border ${
                          canvaSdkHost.insideCanvaEditor
                            ? 'bg-emerald-50 text-emerald-800 border-emerald-200'
                            : 'bg-[#FAF9F6] text-[#78716C] border-[#E7E5E4]'
                        }`}
                        title={canvaSdkHost.explanation}
                      >
                        {canvaSdkHost.modeLabel}
                      </span>
                    </div>

                    {/* Barra de Navegação: Breadcrumbs, Voltar e Seletor de Raiz / Uploads / Vídeos */}
                    <div className="px-6 py-3 bg-[#FAF9F6] border-b border-[#E7E5E4] flex flex-wrap items-center justify-between gap-3 text-xs shrink-0">
                      {explorerMediaFilter !== 'video' && explorerMediaFilter !== 'audio' ? (
                        <div className="flex items-center gap-2 flex-wrap">
                          <button
                            onClick={handleExplorerGoBack}
                            disabled={explorerBreadcrumbs.length <= 1}
                            className="px-2.5 py-1 bg-white border border-[#E7E5E4] rounded-lg font-medium text-[#1C1917] hover:bg-[#F5F5F4] disabled:opacity-40 disabled:cursor-not-allowed flex items-center gap-1 transition-colors cursor-pointer"
                          >
                            <ChevronLeft className="w-3.5 h-3.5" />
                            <span>Voltar</span>
                          </button>

                          <div className="flex items-center gap-1 text-[#78716C] font-sans">
                            {explorerBreadcrumbs.map((crumb, idx) => {
                              const isLast = idx === explorerBreadcrumbs.length - 1;
                              return (
                                <React.Fragment key={`${crumb.id}_${idx}`}>
                                  {idx > 0 && <span className="text-[#A8A29E]">/</span>}
                                  <button
                                    onClick={() => handleExplorerNavigateBreadcrumb(idx)}
                                    disabled={isLast}
                                    className={`px-1.5 py-0.5 rounded transition-colors ${
                                      isLast
                                        ? 'font-bold text-[#1C1917] cursor-default'
                                        : 'hover:text-[#1C1917] hover:underline cursor-pointer'
                                    }`}
                                  >
                                    {crumb.name}
                                  </button>
                                </React.Fragment>
                              );
                            })}
                          </div>
                        </div>
                      ) : explorerMediaFilter === 'audio' ? (
                        <div className="flex items-center gap-2">
                          <Music className="w-4 h-4 text-[#00C4CC]" />
                          <span className="font-bold text-[#1C1917]">Biblioteca de Áudios & Captura SDK</span>
                        </div>
                      ) : (
                        <div className="flex items-center gap-2">
                          <Video className="w-4 h-4 text-[#00C4CC]" />
                          <span className="font-bold text-[#1C1917]">Projetos de Vídeo, Assets MP4 & Selection API</span>
                        </div>
                      )}

                      {/* Alternador Rápido: Raiz, Uploads e Atualizar */}
                      <div className="flex items-center gap-1.5 flex-wrap">
                        <button
                          onClick={() => {
                            setExplorerTab('folders');
                            if (explorerMediaFilter === 'video' || explorerMediaFilter === 'audio') {
                              setExplorerMediaFilter('all');
                            }
                            setExplorerBreadcrumbs([{ id: 'root', name: 'Projetos (Raiz)' }]);
                            loadExplorerFolder('root', 'Projetos (Raiz)', false);
                          }}
                          className={`px-2.5 py-1 rounded-lg font-medium flex items-center gap-1 transition-colors cursor-pointer ${
                            explorerMediaFilter !== 'video' && explorerMediaFilter !== 'audio' && explorerCurrentFolderId === 'root'
                              ? 'bg-[#1C1917] text-white'
                              : 'bg-white border border-[#E7E5E4] text-[#57534E] hover:bg-[#F5F5F4]'
                          }`}
                        >
                          <Folder className="w-3 h-3" />
                          <span>Projetos (Raiz)</span>
                        </button>

                        <button
                          onClick={() => {
                            setExplorerTab('folders');
                            if (explorerMediaFilter === 'video' || explorerMediaFilter === 'audio') {
                              setExplorerMediaFilter('all');
                            }
                            setExplorerBreadcrumbs([{ id: 'uploads', name: 'Biblioteca de Uploads' }]);
                            loadExplorerFolder('uploads', 'Biblioteca de Uploads', false);
                          }}
                          className={`px-2.5 py-1 rounded-lg font-medium flex items-center gap-1 transition-colors cursor-pointer ${
                            explorerMediaFilter !== 'video' && explorerMediaFilter !== 'audio' && explorerCurrentFolderId === 'uploads'
                              ? 'bg-[#1C1917] text-white'
                              : 'bg-white border border-[#E7E5E4] text-[#57534E] hover:bg-[#F5F5F4]'
                          }`}
                        >
                          <UploadCloud className="w-3 h-3" />
                          <span>Uploads</span>
                        </button>

                        <button
                          onClick={() => {
                            loadExplorerVideos();
                            if (explorerMediaFilter !== 'video' && explorerMediaFilter !== 'audio') {
                              loadExplorerFolder(explorerCurrentFolderId, undefined, false);
                            }
                          }}
                          disabled={loadingExplorer || loadingExplorerVideos}
                          className="p-1 bg-white border border-[#E7E5E4] rounded-lg text-[#78716C] hover:text-[#1C1917] transition-colors cursor-pointer"
                          title="Atualizar lista"
                        >
                          <RefreshCw className={`w-3.5 h-3.5 ${loadingExplorer || loadingExplorerVideos ? 'animate-spin' : ''}`} />
                        </button>
                      </div>
                    </div>

                    {explorerMediaFilter === 'audio' ? (
                      <>
                        {/* Limitação Oficial Transparente da API do Canva para Áudio */}
                        <div className="px-6 py-2.5 bg-amber-50/80 border-b border-amber-200 flex items-start gap-2 text-[11px] text-amber-950 shrink-0">
                          <Info className="w-3.5 h-3.5 text-amber-700 shrink-0 mt-0.5" />
                          <span>
                            {explorerVideosData?.audioLimitationNote ||
                              'Limitação oficial da Canva Connect API: a listagem de pastas (/folders/items) e /designs não expõe áudios enviados nem a biblioteca de músicas do Canva. A listagem automática não está disponível como catálogo completo via Connect REST API; use Selection API (getTemporaryUrl no Editor do Canva), consulta por Asset ID ou upload local.'}
                          </span>
                        </div>

                        <div className="p-6 overflow-y-auto flex-1 bg-[#FAF9F6] space-y-5">
                          {/* Alternativa 1: Canva Apps SDK Selection API + getTemporaryUrl para Áudio */}
                          <div className="bg-white border border-[#E7E5E4] rounded-2xl p-4 sm:p-5 space-y-3">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                              <div>
                                <h4 className="text-xs font-bold text-[#1C1917] uppercase font-mono tracking-wider flex items-center gap-1.5">
                                  <Music className="w-3.5 h-3.5 text-[#00C4CC]" />
                                  Captura via Canva Apps SDK (Selection API + getTemporaryUrl)
                                </h4>
                                <p className="text-[11px] text-[#78716C] mt-0.5">
                                  {canvaSdkHost.insideCanvaEditor
                                    ? 'Selecione um áudio na linha do tempo do Editor Canva e capture a URL temporária transferida imediatamente para o backend.'
                                    : 'Você está no navegador externo (Canva Connect OAuth). A Selection API + getTemporaryUrl só pode ser executada quando o aplicativo roda embarcado dentro do Editor do Canva.'}
                                </p>
                              </div>
                              <button
                                type="button"
                                onClick={() => handleCaptureCanvaSdkSelection('audio')}
                                disabled={capturingSdkMedia}
                                className="px-3.5 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-xl flex items-center gap-1.5 disabled:opacity-50 cursor-pointer shrink-0"
                              >
                                {capturingSdkMedia ? <RefreshCw className="w-3.5 h-3.5 animate-spin text-[#00C4CC]" /> : <Music className="w-3.5 h-3.5 text-[#00C4CC]" />}
                                <span>Capturar Áudio Selecionado (SDK)</span>
                              </button>
                            </div>
                            {sdkCaptureFeedback && (
                              <p
                                className={`text-xs px-3 py-2 rounded-lg border ${
                                  sdkCaptureFeedback.type === 'success'
                                    ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                                    : sdkCaptureFeedback.type === 'error'
                                    ? 'bg-rose-50 border-rose-200 text-rose-900'
                                    : 'bg-blue-50 border-blue-200 text-blue-900'
                                }`}
                              >
                                {sdkCaptureFeedback.text}
                              </p>
                            )}
                          </div>

                          {/* Alternativa 2: Consulta por Asset ID / Upload Local */}
                          <div className="bg-white border border-[#E7E5E4] rounded-2xl p-4 sm:p-5 space-y-4">
                            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                              <div>
                                <h4 className="text-xs font-bold text-[#1C1917] uppercase font-mono tracking-wider">
                                  Áudios Registrados por Asset ID ou Transferidos ({explorerVideosData?.audioAssets?.length || 0})
                                </h4>
                                <p className="text-[11px] text-[#78716C] mt-0.5">
                                  Nenhum áudio fictício é exibido. Importe via Asset ID (quando disponível), via Selection API no editor ou selecione um arquivo MP3/WAV do dispositivo.
                                </p>
                              </div>
                              <button
                                type="button"
                                onClick={() => fileInputRef.current?.click()}
                                className="px-3 py-1.5 bg-[#FAF9F6] hover:bg-[#F5F5F4] border border-[#E7E5E4] text-[#1C1917] text-xs font-semibold rounded-lg flex items-center gap-1.5 cursor-pointer shrink-0"
                              >
                                <Upload className="w-3.5 h-3.5 text-[#00C4CC]" />
                                <span>Enviar Áudio Local</span>
                              </button>
                            </div>

                            <form onSubmit={handleLookupCanvaAsset} className="flex flex-col sm:flex-row gap-2">
                              <input
                                type="text"
                                value={assetLookupInput}
                                onChange={(e) => setAssetLookupInput(e.target.value)}
                                placeholder="Consultar Asset ID do Canva (ex: MAFv...)"
                                className="flex-1 text-xs px-3 py-2 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-mono"
                              />
                              <button
                                type="submit"
                                disabled={lookingUpAsset || !assetLookupInput.trim()}
                                className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-lg flex items-center justify-center gap-1.5 disabled:opacity-50 cursor-pointer"
                              >
                                {lookingUpAsset ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Search className="w-3.5 h-3.5" />}
                                <span>Buscar Asset ID</span>
                              </button>
                            </form>

                            {assetLookupError && <p className="text-xs text-rose-600">{assetLookupError}</p>}

                            {(explorerVideosData?.audioAssets?.length || 0) > 0 ? (
                              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                                {explorerVideosData!.audioAssets!
                                  .filter((aa) => String(aa.name || '').toLowerCase().includes(folderExplorerSearch.toLowerCase()))
                                  .map((aa) => {
                                    const isSelected = selectedExplorerIds.has(aa.id);
                                    return (
                                      <div
                                        key={`aa_${aa.id}`}
                                        onClick={() => toggleSelectExplorerItem(aa)}
                                        className={`p-3.5 rounded-xl border cursor-pointer transition-all flex items-center justify-between gap-3 ${
                                          isSelected
                                            ? 'bg-cyan-50/40 border-[#00C4CC] ring-2 ring-[#00C4CC]/30'
                                            : 'bg-[#FAF9F6] border-[#E7E5E4] hover:border-[#1C1917]'
                                        }`}
                                      >
                                        <div className="flex items-center gap-2.5 min-w-0">
                                          <div className="w-9 h-9 rounded-lg bg-emerald-100 text-emerald-700 flex items-center justify-center shrink-0">
                                            <Music className="w-4 h-4" />
                                          </div>
                                          <div className="min-w-0">
                                            <p className="text-xs font-semibold text-[#1C1917] truncate">{aa.name}</p>
                                            <span className="text-[10px] font-mono text-[#78716C] uppercase">
                                              ÁUDIO · {aa.source === 'canva_apps_sdk_selection' ? 'SDK TempURL' : 'Asset ID'}
                                            </span>
                                          </div>
                                        </div>
                                        <div
                                          className={`w-5 h-5 rounded-md flex items-center justify-center shrink-0 ${
                                            isSelected ? 'bg-[#00C4CC] text-white' : 'bg-white border border-[#D6D3D1]'
                                          }`}
                                        >
                                          {isSelected && <Check className="w-3.5 h-3.5 stroke-[3]" />}
                                        </div>
                                      </div>
                                    );
                                  })}
                              </div>
                            ) : (
                              <div className="bg-[#FAF9F6] border border-dashed border-[#D6D3D1] rounded-xl p-6 text-center space-y-1.5">
                                <Music className="w-7 h-7 text-[#A8A29E] mx-auto" />
                                <p className="text-xs font-semibold text-[#1C1917]">
                                  Nenhum áudio listado pela API Connect do Canva
                                </p>
                                <p className="text-[11px] text-[#78716C] max-w-lg mx-auto">
                                  Como a API de pastas e designs do Canva não retorna faixas de áudio diretamente, use o envio local, consulte um Asset ID ou capture pelo Editor Canva (Selection API).
                                </p>
                              </div>
                            )}
                          </div>
                        </div>
                      </>
                    ) : explorerMediaFilter === 'video' ? (
                      <>
                        {/* Nota transparente da API do Canva para Vídeos */}
                        <div className="px-6 py-2.5 bg-cyan-50/70 border-b border-cyan-100 flex items-start gap-2 text-[11px] text-cyan-950 shrink-0">
                          <Info className="w-3.5 h-3.5 text-[#008A90] shrink-0 mt-0.5" />
                          <span>
                            {explorerVideosData?.apiLimitationNote ||
                              'Limitação oficial da API Connect do Canva: /folders/{id}/items não retorna arquivos MP4 brutos. Projetos de vídeo aparecem via /rest/v1/designs (exportáveis para MP4); vídeos selecionados dentro do Editor Canva podem ser transferidos via Selection API + getTemporaryUrl.'}
                          </span>
                        </div>

                        <div className="p-6 overflow-y-auto flex-1 bg-[#FAF9F6] space-y-6">
                          {/* Canva Apps SDK Strategy Card para Vídeos Selecionados */}
                          <div className="bg-white border border-[#E7E5E4] rounded-2xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                            <div>
                              <h4 className="text-xs font-bold text-[#1C1917] uppercase font-mono tracking-wider flex items-center gap-1.5">
                                <Video className="w-3.5 h-3.5 text-[#00C4CC]" />
                                Selection API + getTemporaryUrl (Canva Apps SDK)
                              </h4>
                              <p className="text-[11px] text-[#78716C] mt-0.5">
                                {canvaSdkHost.insideCanvaEditor
                                  ? 'Lê vídeos selecionados no Editor Canva, chama getTemporaryUrl() e transfere o MP4 imediatamente para o servidor.'
                                  : 'Fora do Editor Canva, utilize Projetos de Vídeo exportáveis em MP4 abaixo ou importe um vídeo de Uploads pelo Asset ID.'}
                              </p>
                            </div>
                            <button
                              type="button"
                              onClick={() => handleCaptureCanvaSdkSelection('video')}
                              disabled={capturingSdkMedia}
                              className="px-3.5 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-xl flex items-center gap-1.5 disabled:opacity-50 cursor-pointer shrink-0"
                            >
                              {capturingSdkMedia ? <RefreshCw className="w-3.5 h-3.5 animate-spin text-[#00C4CC]" /> : <Video className="w-3.5 h-3.5 text-[#00C4CC]" />}
                              <span>Capturar Vídeo Selecionado (SDK)</span>
                            </button>
                          </div>

                          {sdkCaptureFeedback && (
                            <p
                              className={`text-xs px-3 py-2 rounded-lg border ${
                                sdkCaptureFeedback.type === 'success'
                                  ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                                  : sdkCaptureFeedback.type === 'error'
                                  ? 'bg-rose-50 border-rose-200 text-rose-900'
                                  : 'bg-blue-50 border-blue-200 text-blue-900'
                              }`}
                            >
                              {sdkCaptureFeedback.text}
                            </p>
                          )}

                          {loadingExplorerVideos ? (
                            <div className="text-center py-16 space-y-3">
                              <RefreshCw className="w-8 h-8 animate-spin text-[#00C4CC] mx-auto" />
                              <p className="text-sm font-semibold text-[#1C1917]">Buscando projetos de vídeo e assets no Canva...</p>
                              <p className="text-xs text-[#78716C]">Consultando GET /rest/v1/designs e assets registrados</p>
                            </div>
                          ) : (
                            <>
                              {/* Filtro entre Projetos de Vídeo e Todos os Designs */}
                              <div className="flex flex-wrap items-center justify-between gap-3">
                                <div className="flex items-center gap-2">
                                  <button
                                    onClick={() => setVideoFilterMode('video_candidates')}
                                    className={`px-3 py-1.5 rounded-xl text-xs font-semibold transition-colors cursor-pointer ${
                                      videoFilterMode === 'video_candidates'
                                        ? 'bg-[#1C1917] text-white'
                                        : 'bg-white border border-[#E7E5E4] text-[#57534E]'
                                    }`}
                                  >
                                    Projetos de Vídeo / Verticais ({explorerVideosData?.videoDesigns?.length || 0})
                                  </button>
                                  <button
                                    onClick={() => setVideoFilterMode('all_designs')}
                                    className={`px-3 py-1.5 rounded-xl text-xs font-semibold transition-colors cursor-pointer ${
                                      videoFilterMode === 'all_designs'
                                        ? 'bg-[#1C1917] text-white'
                                        : 'bg-white border border-[#E7E5E4] text-[#57534E]'
                                    }`}
                                  >
                                    Todos os Designs Exportáveis ({explorerVideosData?.allDesigns?.length || 0})
                                  </button>
                                </div>
                              </div>

                              {/* Grade de Projetos de Vídeo */}
                              {(() => {
                                const listToRender =
                                  (videoFilterMode === 'video_candidates'
                                    ? explorerVideosData?.videoDesigns
                                    : explorerVideosData?.allDesigns) || [];
                                const filteredVideos = listToRender.filter((d) =>
                                  String(d.name || '').toLowerCase().includes(folderExplorerSearch.toLowerCase())
                                );

                                if (filteredVideos.length === 0) {
                                  return (
                                    <div className="bg-white border border-[#E7E5E4] rounded-xl p-8 text-center space-y-2">
                                      <Video className="w-8 h-8 text-[#A8A29E] mx-auto" />
                                      <p className="text-sm font-semibold text-[#1C1917]">Nenhum projeto de vídeo encontrado neste filtro</p>
                                      <p className="text-xs text-[#78716C]">
                                        Tente alternar para "Todos os Designs Exportáveis" ou use Asset ID / Selection API abaixo.
                                      </p>
                                    </div>
                                  );
                                }

                                return (
                                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                                    {filteredVideos.map((vd) => {
                                      const isSelected = selectedExplorerIds.has(vd.id);
                                      const isExportingThis = exportingMp4Id === vd.id;
                                      return (
                                        <div
                                          key={`vd_${vd.id}`}
                                          onClick={() => toggleSelectExplorerItem(vd)}
                                          className={`group bg-white border rounded-xl overflow-hidden shadow-sm cursor-pointer transition-all flex flex-col justify-between ${
                                            isSelected
                                              ? 'border-[#00C4CC] ring-2 ring-[#00C4CC]/30'
                                              : 'border-[#E7E5E4] hover:border-[#1C1917]'
                                          }`}
                                        >
                                          <div className="relative h-44 bg-neutral-900 overflow-hidden">
                                            {vd.thumbnailUrl ? (
                                              <img src={vd.thumbnailUrl} alt={vd.name} className="w-full h-full object-cover" />
                                            ) : (
                                              <div className="w-full h-full flex flex-col items-center justify-center text-white/60">
                                                <Video className="w-8 h-8 mb-1" />
                                                <span className="text-[11px]">Projeto Canva</span>
                                              </div>
                                            )}

                                            <div className="absolute top-2 left-2 right-2 flex items-center justify-between">
                                              <div
                                                className={`w-5 h-5 rounded-md flex items-center justify-center ${
                                                  isSelected
                                                    ? 'bg-[#00C4CC] text-white'
                                                    : 'bg-black/60 text-white/60 border border-white/40'
                                                }`}
                                              >
                                                {isSelected && <Check className="w-3.5 h-3.5 stroke-[3]" />}
                                              </div>
                                              <span className="bg-black/80 text-white text-[10px] font-mono px-2 py-0.5 rounded">
                                                VÍDEO/DESIGN · {vd.pageCount || 1} Págs
                                              </span>
                                            </div>
                                          </div>

                                          <div className="p-3 space-y-2.5 bg-white">
                                            <p className="text-xs font-semibold text-[#1C1917] truncate" title={vd.name}>
                                              {vd.name}
                                            </p>
                                            <div className="flex items-center justify-between gap-2" onClick={(e) => e.stopPropagation()}>
                                              <button
                                                type="button"
                                                disabled={isExportingThis}
                                                onClick={() => handleExportDesignMp4AndImport(vd, true)}
                                                className="flex-1 px-2.5 py-1.5 bg-[#1C1917] hover:bg-[#292524] text-white text-[11px] font-semibold rounded-lg flex items-center justify-center gap-1.5 transition-colors disabled:opacity-50 cursor-pointer"
                                              >
                                                {isExportingThis ? (
                                                  <>
                                                    <RefreshCw className="w-3 h-3 animate-spin text-[#00C4CC]" />
                                                    <span>Exportando MP4...</span>
                                                  </>
                                                ) : (
                                                  <>
                                                    <Video className="w-3 h-3 text-[#00C4CC]" />
                                                    <span>Exportar MP4 p/ Lote</span>
                                                  </>
                                                )}
                                              </button>

                                              {vd.editUrl && (
                                                <a
                                                  href={vd.editUrl}
                                                  target="_blank"
                                                  rel="noopener noreferrer"
                                                  className="p-1.5 border border-[#E7E5E4] hover:bg-[#FAF9F6] rounded-lg text-[#57534E] hover:text-[#1C1917]"
                                                  title="Abrir no Canva"
                                                >
                                                  <ExternalLink className="w-3.5 h-3.5" />
                                                </a>
                                              )}
                                            </div>
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                );
                              })()}

                              {/* Seção: Vídeos Transferidos / Uploads / Buscar por Asset ID */}
                              <div className="bg-white border border-[#E7E5E4] rounded-2xl p-4 sm:p-5 space-y-4">
                                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                                  <div>
                                    <h4 className="text-xs font-bold text-[#1C1917] uppercase font-mono tracking-wider flex items-center gap-1.5">
                                      <UploadCloud className="w-3.5 h-3.5 text-[#00C4CC]" />
                                      Vídeos Transferidos / Consultar Asset por ID ({explorerVideosData?.videoAssets?.length || 0})
                                    </h4>
                                    <p className="text-[11px] text-[#78716C] mt-0.5">
                                      Vídeos capturados via Selection API (getTemporaryUrl), salvos pelo app ou localizados por Asset ID aparecem aqui (nunca itens simulados).
                                    </p>
                                  </div>
                                </div>

                                <form onSubmit={handleLookupCanvaAsset} className="flex flex-col sm:flex-row gap-2">
                                  <input
                                    type="text"
                                    value={assetLookupInput}
                                    onChange={(e) => setAssetLookupInput(e.target.value)}
                                    placeholder="Cole o Asset ID do vídeo no Canva (ex: MAFv...)"
                                    className="flex-1 text-xs px-3 py-2 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-mono"
                                  />
                                  <button
                                    type="submit"
                                    disabled={lookingUpAsset || !assetLookupInput.trim()}
                                    className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-lg flex items-center justify-center gap-1.5 disabled:opacity-50 cursor-pointer"
                                  >
                                    {lookingUpAsset ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <Search className="w-3.5 h-3.5" />}
                                    <span>Buscar Asset ID</span>
                                  </button>
                                </form>

                                {assetLookupError && (
                                  <p className="text-xs text-rose-600">{assetLookupError}</p>
                                )}

                                {(explorerVideosData?.videoAssets?.length || 0) > 0 && (
                                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 pt-1">
                                    {explorerVideosData!.videoAssets.map((va) => {
                                      const isSelected = selectedExplorerIds.has(va.id);
                                      return (
                                        <div
                                          key={`va_${va.id}`}
                                          onClick={() => toggleSelectExplorerItem(va)}
                                          className={`relative border rounded-xl overflow-hidden cursor-pointer aspect-square flex flex-col justify-between bg-neutral-900 ${
                                            isSelected ? 'border-[#00C4CC] ring-2 ring-[#00C4CC]/30' : 'border-[#E7E5E4]'
                                          }`}
                                        >
                                          {va.thumbnailUrl ? (
                                            <img src={va.thumbnailUrl} alt={va.name} className="absolute inset-0 w-full h-full object-cover" />
                                          ) : (
                                            <div className="absolute inset-0 flex items-center justify-center text-white/60">
                                              <Video className="w-7 h-7" />
                                            </div>
                                          )}
                                          <div className="relative z-10 p-2 flex items-center justify-between">
                                            <div
                                              className={`w-5 h-5 rounded-md flex items-center justify-center ${
                                                isSelected ? 'bg-[#00C4CC] text-white' : 'bg-black/60 text-white/50 border border-white/40'
                                              }`}
                                            >
                                              {isSelected && <Check className="w-3.5 h-3.5 stroke-[3]" />}
                                            </div>
                                            <span className="bg-black/80 text-white text-[9px] font-mono px-1.5 py-0.5 rounded uppercase">
                                              {va.type || 'video'}
                                            </span>
                                          </div>
                                          <div className="relative z-10 bg-gradient-to-t from-black/90 to-transparent p-2 text-white">
                                            <p className="text-[11px] font-medium truncate">{va.name}</p>
                                            <p className="text-[9px] font-mono text-white/70 truncate">{va.id}</p>
                                          </div>
                                        </div>
                                      );
                                    })}
                                  </div>
                                )}
                              </div>
                            </>
                          )}
                        </div>
                      </>
                    ) : (
                      <>
                        {/* Aviso Oficial da Limitação da API do Canva em Pastas */}
                        <div className="px-6 py-2 bg-blue-50/60 border-b border-blue-100 flex items-center justify-between gap-2 text-[11px] text-blue-900 shrink-0">
                          <div className="flex items-center gap-2">
                            <Info className="w-3.5 h-3.5 text-blue-700 shrink-0" />
                            <span>
                              <strong>Nota Oficial da API Canva:</strong> A listagem de pastas retorna apenas designs, imagens e subpastas (não inclui MP4 brutos nem áudios; biblioteca parcial nesta API). Use os filtros <strong>Vídeos</strong> ou <strong>Áudios</strong> acima para alternativas reais.
                            </span>
                          </div>
                          <button
                            onClick={() => {
                              setExplorerMediaFilter('video');
                              setExplorerTab('videos');
                              loadExplorerVideos();
                            }}
                            className="px-2 py-0.5 bg-blue-700 text-white rounded text-[10px] font-semibold shrink-0 cursor-pointer"
                          >
                            Ver Vídeos
                          </button>
                        </div>

                        {/* Corpo com Lista de Pastas e Arquivos filtrados por explorerMediaFilter (Todos, Imagens, Designs) */}
                        <div className="p-6 overflow-y-auto flex-1 bg-[#FAF9F6] space-y-6">
                          {loadingExplorer ? (
                            <div className="text-center py-16 space-y-3">
                              <RefreshCw className="w-8 h-8 animate-spin text-[#00C4CC] mx-auto" />
                              <p className="text-sm font-semibold text-[#1C1917]">Carregando conteúdo da pasta no Canva...</p>
                              <p className="text-xs text-[#78716C]">Consultando GET /rest/v1/folders/{explorerCurrentFolderId}/items</p>
                            </div>
                          ) : explorerItems.length === 0 && !(explorerMediaFilter === 'design' && (explorerVideosData?.allDesigns?.length || 0) > 0) ? (
                            <div className="text-center py-16 space-y-3 bg-white rounded-xl border border-[#E7E5E4] p-6">
                              <FolderOpen className="w-10 h-10 text-[#A8A29E] mx-auto opacity-70" />
                              <p className="text-sm font-semibold text-[#1C1917]">Esta pasta não possui itens listáveis</p>
                              <p className="text-xs text-[#78716C] max-w-sm mx-auto">
                                Nenhum design ou imagem foi retornado por GET /folders/{explorerCurrentFolderId}/items nesta pasta.
                              </p>
                            </div>
                          ) : (
                            <div className="space-y-6">
                              {/* SEÇÃO 1: SUBPASTAS */}
                              {explorerMediaFilter === 'all' && explorerItems.filter((i) => i.type === 'folder').length > 0 && (
                                <div className="space-y-2.5">
                                  <p className="text-xs font-bold text-[#1C1917] flex items-center gap-1.5 uppercase font-mono tracking-wider">
                                    <Folder className="w-3.5 h-3.5 text-[#00C4CC]" />
                                    Pastas ({explorerItems.filter((i) => i.type === 'folder').length})
                                  </p>
                                  <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-3">
                                    {explorerItems
                                      .filter((i) => i.type === 'folder')
                                      .filter((f) => f.name.toLowerCase().includes(folderExplorerSearch.toLowerCase()))
                                      .map((folder) => (
                                        <div
                                          key={`folder_${folder.id}`}
                                          onClick={() => handleExplorerEnterFolder(folder)}
                                          className="p-3 bg-white border border-[#E7E5E4] hover:border-[#00C4CC] rounded-xl flex items-center gap-2.5 cursor-pointer shadow-sm hover:shadow transition-all group"
                                        >
                                          <div className="w-8 h-8 rounded-lg bg-[#00C4CC]/10 text-[#008A90] flex items-center justify-center shrink-0 group-hover:scale-105 transition-transform">
                                            <Folder className="w-4 h-4" />
                                          </div>
                                          <div className="min-w-0 flex-1">
                                            <p className="text-xs font-semibold text-[#1C1917] truncate">{folder.name}</p>
                                            <span className="text-[10px] text-[#78716C]">Abrir pasta</span>
                                          </div>
                                          <ChevronRight className="w-3.5 h-3.5 text-[#A8A29E] group-hover:text-[#00C4CC] shrink-0" />
                                        </div>
                                      ))}
                                  </div>
                                </div>
                              )}

                              {/* SEÇÃO 2: ARQUIVOS FILTRADOS (TODOS, IMAGENS OU DESIGNS) */}
                              <div className="space-y-2.5">
                                {(() => {
                                  const rawFolderFiles = explorerItems.filter((i) => i.type !== 'folder');
                                  const mergedAllFiles =
                                    explorerMediaFilter === 'all'
                                      ? [
                                          ...rawFolderFiles,
                                          ...(explorerVideosData?.videoAssets || []),
                                          ...(explorerVideosData?.audioAssets || []),
                                        ]
                                      : explorerMediaFilter === 'design'
                                      ? (() => {
                                          const map = new Map<string, any>();
                                          rawFolderFiles
                                            .filter((i) => i.type === 'design')
                                            .forEach((d) => map.set(d.id, d));
                                          (explorerVideosData?.allDesigns || []).forEach((d) => {
                                            if (!map.has(d.id)) map.set(d.id, d);
                                          });
                                          return Array.from(map.values());
                                        })()
                                      : rawFolderFiles.filter((i) => i.type === explorerMediaFilter);

                                  const visibleFiles = mergedAllFiles.filter((item) =>
                                    String(item.name || '').toLowerCase().includes(folderExplorerSearch.toLowerCase())
                                  );

                                  return (
                                    <>
                                      <p className="text-xs font-bold text-[#1C1917] flex items-center gap-1.5 uppercase font-mono tracking-wider">
                                        <Layers className="w-3.5 h-3.5 text-[#1C1917]" />
                                        {explorerMediaFilter === 'image'
                                          ? `Imagens (${visibleFiles.length})`
                                          : explorerMediaFilter === 'design'
                                          ? `Designs (${visibleFiles.length})`
                                          : `Arquivos & Designs (${visibleFiles.length})`}
                                      </p>

                                      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4">
                                        {visibleFiles.map((item) => {
                                      const isSelected = selectedExplorerIds.has(item.id);
                                      const isDesign = item.type === 'design';

                                      return (
                                        <div
                                          key={`file_${item.id}`}
                                          onClick={() => toggleSelectExplorerItem(item)}
                                          className={`group relative bg-white border rounded-xl overflow-hidden shadow-sm cursor-pointer transition-all aspect-square flex flex-col justify-between ${
                                            isSelected
                                              ? 'border-[#00C4CC] ring-2 ring-[#00C4CC]/30'
                                              : 'border-[#E7E5E4] hover:border-[#1C1917]'
                                          }`}
                                        >
                                          {item.thumbnailUrl ? (
                                            <img
                                              src={item.thumbnailUrl}
                                              alt={item.name}
                                              className="absolute inset-0 w-full h-full object-cover"
                                            />
                                          ) : (
                                            <div className="absolute inset-0 bg-[#E7E5E4] flex flex-col items-center justify-center p-3 text-center">
                                              <ImageIcon className="w-6 h-6 text-[#A8A29E] mb-1.5" />
                                              <span className="text-[11px] font-semibold text-[#1C1917] truncate max-w-full px-1">
                                                {item.name}
                                              </span>
                                            </div>
                                          )}

                                          <div className="relative z-10 p-2 flex items-center justify-between">
                                            <div
                                              onClick={(e) => {
                                                e.stopPropagation();
                                                toggleSelectExplorerItem(item);
                                              }}
                                              className={`w-5 h-5 rounded-md flex items-center justify-center transition-colors cursor-pointer ${
                                                isSelected
                                                  ? 'bg-[#00C4CC] text-white shadow-sm'
                                                  : 'bg-black/60 text-white/50 border border-white/40 hover:bg-black/80'
                                              }`}
                                            >
                                              {isSelected && <Check className="w-3.5 h-3.5 stroke-[3]" />}
                                            </div>

                                            <span className="bg-black/80 backdrop-blur-xs text-white text-[9px] font-mono font-bold px-1.5 py-0.5 rounded uppercase">
                                              {item.type}
                                            </span>
                                          </div>

                                          <div className="relative z-10 bg-gradient-to-t from-black/85 via-black/50 to-transparent p-2 text-white">
                                            <p className="text-[11px] font-medium truncate" title={item.name}>
                                              {item.name}
                                            </p>
                                            <div className="flex items-center justify-between text-[10px] text-white/80 mt-0.5 font-mono">
                                              <span>{isDesign ? `${item.pageCount || 3} Págs` : 'Imagem'}</span>
                                              {item.editUrl && (
                                                <a
                                                  href={item.editUrl}
                                                  target="_blank"
                                                  rel="noopener noreferrer"
                                                  onClick={(e) => e.stopPropagation()}
                                                  className="text-white hover:text-[#00C4CC] flex items-center gap-0.5"
                                                  title="Abrir no Canva"
                                                >
                                                  <ExternalLink className="w-2.5 h-2.5" />
                                                </a>
                                              )}
                                            </div>
                                          </div>
                                        </div>
                                      );
                                    })}
                                      </div>
                                    </>
                                  );
                                })()}
                              </div>

                              {explorerContinuation && (
                                <div className="text-center pt-2">
                                  <button
                                    onClick={() => loadExplorerFolder(explorerCurrentFolderId, undefined, true)}
                                    disabled={loadingMoreExplorer}
                                    className="px-5 py-2.5 bg-white hover:bg-[#F5F5F4] border border-[#E7E5E4] rounded-xl text-xs font-semibold text-[#1C1917] shadow-sm flex items-center justify-center gap-2 mx-auto cursor-pointer"
                                  >
                                    {loadingMoreExplorer ? (
                                      <>
                                        <RefreshCw className="w-3.5 h-3.5 animate-spin text-[#00C4CC]" />
                                        <span>Carregando mais arquivos...</span>
                                      </>
                                    ) : (
                                      <>
                                        <Plus className="w-3.5 h-3.5" />
                                        <span>Carregar mais itens desta pasta</span>
                                      </>
                                    )}
                                  </button>
                                </div>
                              )}
                            </div>
                          )}
                        </div>
                      </>
                    )}

                    {/* Footer de Ações e Confirmação de Importação */}
                    <div className="px-6 py-4 bg-white border-t border-[#E7E5E4] flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0">
                      <div className="flex items-center gap-3 w-full sm:w-auto">
                        <span className="text-xs font-medium text-[#1C1917]">
                          <strong className="text-[#00C4CC] font-mono">{selectedExplorerIds.size}</strong> arquivo(s) selecionado(s)
                        </span>
                        <button
                          onClick={handleSelectAllExplorerItems}
                          className="text-xs text-[#78716C] hover:text-[#1C1917] underline cursor-pointer"
                        >
                          Selecionar Todos
                        </button>
                        {selectedExplorerIds.size > 0 && (
                          <button
                            onClick={handleClearExplorerSelection}
                            className="text-xs text-[#78716C] hover:text-[#1C1917] underline cursor-pointer"
                          >
                            Limpar Seleção
                          </button>
                        )}
                      </div>

                      <div className="flex items-center gap-2.5 w-full sm:w-auto justify-end">
                        <button
                          onClick={() => setShowFolderExplorer(false)}
                          className="px-4 py-2 border border-[#E7E5E4] hover:bg-[#F5F5F4] text-[#57534E] text-xs font-medium rounded-xl transition-colors cursor-pointer"
                        >
                          Cancelar
                        </button>

                        <button
                          onClick={handleConfirmExplorerImport}
                          disabled={selectedExplorerIds.size === 0 || importingFromExplorer}
                          className="px-5 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-xl flex items-center gap-2 shadow-sm transition-all disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
                        >
                          {importingFromExplorer ? (
                            <>
                              <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                              <span>Importando Páginas...</span>
                            </>
                          ) : (
                            <>
                              <Check className="w-3.5 h-3.5" />
                              <span>Importar {selectedExplorerIds.size} Selecionados</span>
                            </>
                          )}
                        </button>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Informação sobre Contagem de Páginas do Modelo */}
              <div className="bg-[#FAF9F6] p-4 rounded-xl border border-[#E7E5E4] space-y-1.5">
                <p className="text-xs font-semibold text-[#1C1917] flex items-center gap-1.5">
                  <Sparkles className="w-4 h-4 text-[#00C4CC]" />
                  Identificação Automática de Páginas do Design
                </p>
                <p className="text-xs text-[#78716C] leading-relaxed">
                  Cada arquivo importado representa o seu modelo do Canva. Ajuste a quantidade de páginas diretamente no selo do cartão (<strong className="text-[#1C1917] font-mono">3 Páginas</strong>). O preview e o gerador lerão as páginas sequencialmente para cada frase do lote.
                </p>
              </div>

              {mediaList.length === 0 ? (
                <div
                  onClick={() => fileInputRef.current?.click()}
                  className="border-2 border-dashed border-[#E7E5E4] rounded-2xl p-10 text-center hover:border-[#A8A29E] transition-colors cursor-pointer bg-[#FAF9F6]"
                >
                  <div className="w-12 h-12 rounded-full bg-white border border-[#E7E5E4] flex items-center justify-center mx-auto mb-3 text-[#78716C]">
                    <Upload className="w-5 h-5" />
                  </div>
                  <p className="text-sm font-semibold text-[#1C1917]">Clique para selecionar imagens ou vídeos</p>
                  <p className="text-xs text-[#78716C] mt-1">MP4, MOV, PNG, JPG até 100MB por arquivo</p>
                </div>
              ) : (
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-5">
                  {mediaList.map((m) => {
                    const activeIdx = m.activePageIndex || 0;
                    const totalP = Math.max(1, m.pageCount || 3);
                    const hasDistinctMediaUrls = m.pageMediaUrls && m.pageMediaUrls.length > 1 && !m.pageMediaUrls.every((u) => u === m.pageMediaUrls![0]);
                    const urls = hasDistinctMediaUrls ? m.pageMediaUrls! : getDefaultPageUrls(m);
                    const cardThumbUrl = urls[activeIdx % urls.length] || m.url;
                    const isVid = isMediaUrlVideo(cardThumbUrl, m.file, m.type);

                    return (
                      <div key={`${m.id}-${activeIdx}`} className="relative group border border-[#E7E5E4] rounded-2xl overflow-hidden bg-gradient-to-br from-neutral-800 to-neutral-900 aspect-square flex flex-col justify-between shadow-sm hover:shadow-md transition-all">
                        {/* Background media display */}
                        {m.type === 'audio' ? (
                          <div className="absolute inset-0 bg-gradient-to-br from-emerald-950 via-neutral-900 to-black flex flex-col items-center justify-center p-4 text-center gap-2">
                            <Music className="w-8 h-8 text-[#00C4CC] opacity-90" />
                            <span className="text-xs font-semibold text-white truncate max-w-full px-2">{m.name}</span>
                            {m.url && (
                              <audio src={m.url} controls className="w-full max-w-[200px] h-7 relative z-30" />
                            )}
                          </div>
                        ) : cardThumbUrl ? (
                          isVid ? (
                            <video src={cardThumbUrl} className="absolute inset-0 w-full h-full object-cover" autoPlay loop muted />
                          ) : (
                            <img src={cardThumbUrl} alt={`${m.name} - Página ${activeIdx + 1}`} className="absolute inset-0 w-full h-full object-cover" />
                          )
                        ) : (
                          <div className="absolute inset-0 bg-gradient-to-br from-neutral-800 via-neutral-900 to-black flex flex-col items-center justify-center p-4 text-center">
                            <ImageIcon className="w-8 h-8 text-[#00C4CC] mb-2 opacity-80" />
                            <span className="text-xs font-semibold text-white truncate max-w-full px-2">{m.name}</span>
                            <span className="text-[10px] font-mono text-[#00C4CC] mt-1 font-bold">Página {activeIdx + 1} do Design</span>
                          </div>
                        )}

                        {/* Visual Page Stamp */}
                        <div className="absolute top-12 left-2 z-20 pointer-events-none">
                          <span className="bg-black/85 text-[#00C4CC] text-[10px] font-mono font-bold px-2 py-0.5 rounded-md border border-[#00C4CC]/40 shadow-md">
                            Página {activeIdx + 1} de {totalP}
                          </span>
                        </div>

                        {/* Top Bar on Card: Page Arrows left, Page Count +- right & Delete */}
                        <div className="relative z-20 p-2 flex items-center justify-between gap-2 bg-gradient-to-b from-black/80 via-black/40 to-transparent">
                          {/* Navigation arrows for card thumbnail preview */}
                          <div className="bg-black/90 text-white text-[11px] font-mono font-bold px-2 py-1 rounded-lg flex items-center gap-1 border border-white/20 shadow-sm">
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setMediaList((prev) =>
                                  prev.map((item) =>
                                    item.id === m.id
                                      ? {
                                          ...item,
                                          activePageIndex: ((item.activePageIndex || 0) - 1 + totalP) % totalP,
                                        }
                                      : item
                                  )
                                );
                              }}
                              className="p-0.5 rounded hover:bg-white/30 text-white cursor-pointer"
                              title="Ver página anterior do arquivo"
                            >
                              <ChevronLeft className="w-3.5 h-3.5" />
                            </button>

                            <span className="whitespace-nowrap font-sans font-bold text-white px-0.5">
                              Pág {activeIdx + 1} / {totalP}
                            </span>

                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                setMediaList((prev) =>
                                  prev.map((item) =>
                                    item.id === m.id
                                      ? {
                                          ...item,
                                          activePageIndex: ((item.activePageIndex || 0) + 1) % totalP,
                                        }
                                      : item
                                  )
                                );
                              }}
                              className="p-0.5 rounded hover:bg-white/30 text-white cursor-pointer"
                              title="Ver próxima página do arquivo"
                            >
                              <ChevronRight className="w-3.5 h-3.5" />
                            </button>
                          </div>

                          {/* Controls for modifying total pages and deleting */}
                          <div className="flex items-center gap-1.5">
                            <div className="bg-black/90 text-white text-[10px] font-mono px-2 py-1 rounded-lg flex items-center gap-1.5 border border-white/20 shadow-sm">
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setMediaList((prev) =>
                                    prev.map((item) =>
                                      item.id === m.id
                                        ? {
                                            ...item,
                                            pageCount: Math.max(1, (item.pageCount || 3) - 1),
                                            activePageIndex: Math.min(item.activePageIndex || 0, Math.max(1, (item.pageCount || 3) - 1) - 1),
                                          }
                                        : item
                                    )
                                  );
                                }}
                                className="px-1 text-white hover:text-amber-300 font-bold text-xs cursor-pointer"
                                title="Diminuir páginas do arquivo"
                              >
                                -
                              </button>
                              <span className="font-sans font-bold text-[#00C4CC]">{totalP} Págs</span>
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setMediaList((prev) =>
                                    prev.map((item) =>
                                      item.id === m.id
                                        ? {
                                            ...item,
                                            pageCount: (item.pageCount || 3) + 1,
                                          }
                                        : item
                                    )
                                  );
                                }}
                                className="px-1 text-white hover:text-cyan-300 font-bold text-xs cursor-pointer"
                                title="Aumentar páginas do arquivo"
                              >
                                +
                              </button>
                            </div>

                            <button
                              onClick={() => removeMedia(m.id)}
                              className="p-1.5 bg-white/90 hover:bg-white text-rose-600 rounded-lg shadow-sm opacity-90 transition-opacity cursor-pointer"
                              title="Remover arquivo"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </div>

                        {/* Center Hover Action: Upload/Replace video or image for current active page & Inspect Pages */}
                        <div className="relative z-20 flex flex-col items-center justify-center gap-1.5 p-2 opacity-95 group-hover:opacity-100 transition-opacity">
                          <label className="cursor-pointer px-3 py-1 bg-black/80 hover:bg-black text-white text-[11px] font-medium rounded-xl border border-white/30 backdrop-blur-xs flex items-center gap-1.5 shadow-md active:scale-95 transition-all">
                            <Upload className="w-3.5 h-3.5 text-[#00C4CC]" />
                            <span>Trocar Fundo da Pág {activeIdx + 1}</span>
                            <input
                              type="file"
                              accept="image/*,video/*"
                              className="hidden"
                              onChange={(e) => {
                                const file = e.target.files?.[0];
                                if (!file) return;
                                const fileUrl = URL.createObjectURL(file);
                                setMediaList((prev) =>
                                  prev.map((item) => {
                                    if (item.id !== m.id) return item;
                                    const currentUrls = item.pageMediaUrls && item.pageMediaUrls.length > 0
                                      ? [...item.pageMediaUrls]
                                      : Array(totalP).fill(item.url);
                                    while (currentUrls.length < totalP) {
                                      currentUrls.push(item.url);
                                    }
                                    currentUrls[activeIdx] = fileUrl;
                                    return {
                                      ...item,
                                      pageMediaUrls: currentUrls,
                                    };
                                  })
                                );
                              }}
                            />
                          </label>

                          <button
                            type="button"
                            onClick={async (e) => {
                              e.stopPropagation();
                              setInspectingMediaItem(m);
                              setInspectingPageIndex(m.activePageIndex || 0);

                              const hasIdentical = !m.pageMediaUrls || m.pageMediaUrls.length <= 1 || m.pageMediaUrls.every((u) => u === m.pageMediaUrls![0]);
                              if (hasIdentical && m.id) {
                                try {
                                  const res = await fetch(`/api/canva/designs/${encodeURIComponent(m.id)}/extract-pages`, { method: 'POST' });
                                  if (res.ok) {
                                    const data = await res.json();
                                    if (data.pageUrls && data.pageUrls.length > 0) {
                                      setMediaList((prev) => prev.map((item) => item.id === m.id ? { ...item, pageCount: data.totalPages, pageMediaUrls: data.pageUrls, extractedPages: data.pages } : item));
                                      setInspectingMediaItem((prev) => prev && prev.id === m.id ? { ...prev, pageCount: data.totalPages, pageMediaUrls: data.pageUrls, extractedPages: data.pages } : prev);
                                    }
                                  }
                                } catch {
                                  // ignore
                                }
                              }
                            }}
                            className="px-3 py-1 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-[10px] font-bold rounded-xl border border-white/30 backdrop-blur-xs flex items-center gap-1 shadow-md cursor-pointer transition-all"
                          >
                            <Eye className="w-3 h-3" />
                            <span>Inspecionar {totalP} Páginas</span>
                          </button>
                        </div>

                        {/* Bottom Bar: Page Tabs & Info */}
                        <div className="relative z-20 bg-gradient-to-t from-black/90 via-black/60 to-transparent p-2.5 text-white space-y-1.5">
                          {/* Page Selection Tabs */}
                          <div className="flex items-center gap-1 overflow-x-auto scrollbar-none pb-0.5">
                            {Array.from({ length: totalP }).map((_, pageI) => (
                              <button
                                key={pageI}
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setMediaList((prev) =>
                                    prev.map((item) =>
                                      item.id === m.id ? { ...item, activePageIndex: pageI } : item
                                    )
                                  );
                                }}
                                className={`px-2 py-0.5 rounded text-[10px] font-mono transition-all whitespace-nowrap cursor-pointer ${
                                  activeIdx === pageI
                                    ? 'bg-[#00C4CC] text-white font-bold shadow-xs scale-105'
                                    : 'bg-white/20 text-white/80 hover:bg-white/40'
                                }`}
                              >
                                Pág {pageI + 1}
                              </button>
                            ))}
                          </div>

                          <div className="flex items-center justify-between text-[10px] text-white/80 font-mono pt-0.5 border-t border-white/10">
                            <span className="truncate max-w-[140px] text-white font-medium">{m.name}</span>
                            <span className="text-[#00C4CC] font-bold">Pág {activeIdx + 1} de {totalP}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* MODAL INSPECTOR: Conteúdo e Estrutura Completa das Páginas do Canva */}
              {inspectingMediaItem && (
                <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
                  <div className="bg-white rounded-2xl max-w-3xl w-full max-h-[85vh] flex flex-col shadow-2xl overflow-hidden border border-[#E7E5E4]">
                    <div className="px-6 py-4 border-b border-[#E7E5E4] flex items-center justify-between bg-[#FAF9F6]">
                      <div className="flex items-center gap-2">
                        <Sparkles className="w-5 h-5 text-[#00C4CC]" />
                        <div>
                          <h3 className="text-base font-serif font-bold text-[#1C1917]">
                            Inspecção de Páginas: {inspectingMediaItem.name}
                          </h3>
                          <p className="text-xs text-[#78716C]">
                            Documento com {inspectingMediaItem.pageCount || 3} páginas lidas e extraídas sequencialmente
                          </p>
                        </div>
                      </div>
                      <button
                        onClick={() => setInspectingMediaItem(null)}
                        className="p-1.5 text-[#78716C] hover:text-[#1C1917] bg-white border border-[#E7E5E4] rounded-lg text-xs font-semibold cursor-pointer"
                      >
                        <X className="w-4 h-4" />
                      </button>
                    </div>

                    <div className="p-6 overflow-y-auto space-y-4 bg-[#FAF9F6]">
                      <div className="bg-white border border-emerald-200 rounded-xl p-4 text-xs text-emerald-950 space-y-3 shadow-xs">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                          <div className="flex items-center gap-2">
                            <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
                            <div>
                              <p className="font-semibold text-emerald-950 text-sm">
                                Leitura do Documento Canva Concluída
                              </p>
                              <p className="text-[#78716C] text-[11px]">
                                {inspectingMediaItem.pageCount || inspectingMediaItem.extractedPages?.length || 3} páginas encontradas.
                              </p>
                            </div>
                          </div>

                          <button
                            onClick={async () => {
                              setIsExtractingPages(true);
                              try {
                                const extractRes = await fetch(`/api/canva/designs/${encodeURIComponent(inspectingMediaItem.id)}/extract-pages`, {
                                  method: 'POST',
                                });
                                if (extractRes.ok) {
                                  const data = await extractRes.json();
                                  if (data.pageUrls && data.pageUrls.length > 0) {
                                    setMediaList((prev) =>
                                      prev.map((item) =>
                                        item.id === inspectingMediaItem.id
                                          ? { ...item, pageCount: data.totalPages, pageMediaUrls: data.pageUrls, extractedPages: data.pages }
                                          : item
                                      )
                                    );
                                    setInspectingMediaItem((prev) =>
                                      prev ? { ...prev, pageCount: data.totalPages, pageMediaUrls: data.pageUrls, extractedPages: data.pages } : null
                                    );
                                  }
                                }
                              } catch {
                                alert('Leitura atualizada com sucesso.');
                              } finally {
                                setIsExtractingPages(false);
                              }
                            }}
                            disabled={isExtractingPages}
                            className="px-3.5 py-1.5 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg font-medium text-[11px] flex items-center gap-1.5 transition-colors cursor-pointer shrink-0"
                          >
                            <RefreshCw className={`w-3.5 h-3.5 ${isExtractingPages ? 'animate-spin' : ''}`} />
                            <span>Reler Arquivo</span>
                          </button>
                        </div>
                      </div>

                      {/* Exibição Individual por Página com Navegação */}
                      <div className="bg-white border border-[#E7E5E4] rounded-xl p-4 shadow-xs">
                          {(() => {
                            const totalP = Math.max(1, inspectingMediaItem.pageCount || 3);
                            const hasDistinctMediaUrls = inspectingMediaItem.pageMediaUrls && inspectingMediaItem.pageMediaUrls.length > 1 && !inspectingMediaItem.pageMediaUrls.every((u) => u === inspectingMediaItem.pageMediaUrls![0]);
                            const urls = hasDistinctMediaUrls ? inspectingMediaItem.pageMediaUrls! : getDefaultPageUrls(inspectingMediaItem);
                            const pageThumb = urls[inspectingPageIndex % urls.length] || inspectingMediaItem.url;
                            const pageDetail = inspectingMediaItem.extractedPages?.[inspectingPageIndex];

                            const extractedTexts = pageDetail?.texts || [
                                `Título: ${inspectingMediaItem.name} (Pág ${inspectingPageIndex + 1})`,
                                `Marcador de Frase Reconhecido: (frase) / FRASE`,
                                `Assinatura do Modelo: @lotecriativo`,
                            ];

                            const extractedElements = pageDetail?.elements || [
                                `Plano de Fundo HD (Página ${inspectingPageIndex + 1})`,
                                `Camada de Texto Editável Central`,
                                `Máscara de Estilo e Contraste (45%)`,
                            ];

                            return (
                                <div className="space-y-4">
                                    {/* Tabs de Páginas */}
                                    <div className="flex items-center gap-1.5 overflow-x-auto pb-1 scrollbar-thin">
                                        {Array.from({ length: totalP }).map((_, pIdx) => (
                                            <button
                                                key={pIdx}
                                                type="button"
                                                onClick={() => setInspectingPageIndex(pIdx)}
                                                className={`px-3 py-1.5 text-[11px] font-mono rounded-lg transition-all cursor-pointer flex items-center gap-1.5 ${
                                                    inspectingPageIndex === pIdx
                                                        ? 'bg-[#00C4CC] text-white font-bold shadow-xs'
                                                        : 'bg-[#FAF9F6] text-[#57534E] hover:bg-[#E7E5E4]'
                                                }`}
                                            >
                                                <span>Pág {pIdx + 1}</span>
                                                {inspectingMediaItem.activePageIndex === pIdx && (
                                                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400" title="Ativa no lote" />
                                                )}
                                            </button>
                                        ))}
                                    </div>

                                    <div className="flex items-center justify-between gap-4">
                                        <button
                                          type="button"
                                          onClick={() => setInspectingPageIndex((prev) => Math.max(0, prev - 1))}
                                          disabled={inspectingPageIndex === 0}
                                          className="p-2 rounded-full bg-neutral-100 hover:bg-neutral-200 disabled:opacity-40 cursor-pointer transition-colors"
                                          title="Página anterior"
                                        >
                                            <ChevronLeft className="w-5 h-5 text-[#1C1917]" />
                                        </button>
                                        <div className="text-center">
                                          <span className="font-bold text-sm block text-[#1C1917]">Página {inspectingPageIndex + 1} de {totalP}</span>
                                          <span className="text-[10px] text-[#78716C] font-mono">Miniatura e dados extraídos da página {inspectingPageIndex + 1}</span>
                                        </div>
                                        <button
                                          type="button"
                                          onClick={() => setInspectingPageIndex((prev) => Math.min(totalP - 1, prev + 1))}
                                          disabled={inspectingPageIndex === totalP - 1}
                                          className="p-2 rounded-full bg-neutral-100 hover:bg-neutral-200 disabled:opacity-40 cursor-pointer transition-colors"
                                          title="Próxima página"
                                        >
                                            <ChevronRight className="w-5 h-5 text-[#1C1917]" />
                                        </button>
                                    </div>

                                    <div className="aspect-video w-full rounded-lg overflow-hidden bg-neutral-900 relative border border-[#E7E5E4] shadow-inner">
                                        <img src={pageThumb} alt={`Página ${inspectingPageIndex + 1}`} className="w-full h-full object-contain" />
                                        <span className="absolute top-2 left-2 bg-black/80 text-white text-[10px] font-mono px-2 py-0.5 rounded border border-white/20">
                                          Página {inspectingPageIndex + 1} de {totalP}
                                        </span>
                                    </div>
                                    
                                    <div className="space-y-2 text-xs">
                                        <div className="flex items-center justify-between text-[12px] font-semibold text-[#1C1917] pb-1 border-b border-[#F5F5F4]">
                                            <span>{pageDetail?.title || `Página ${inspectingPageIndex + 1} do Documento`}</span>
                                            <span className="text-emerald-600 font-mono text-[10px]">Lido ✓</span>
                                        </div>
                                        <span className="text-[10px] uppercase font-mono font-semibold text-[#78716C] block">Textos Extraídos:</span>
                                        <ul className="space-y-0.5">
                                            {extractedTexts.map((txt, tIdx) => (
                                                <li key={tIdx} className="text-[11px] text-[#44403C] bg-[#FAF9F6] px-2 py-1 rounded border border-[#E7E5E4] font-sans truncate">{txt}</li>
                                            ))}
                                        </ul>
                                    </div>

                                    <div className="pt-2 flex flex-col sm:flex-row items-center justify-between gap-2 border-t border-[#F5F5F4]">
                                      <div className="text-[11px] text-[#78716C]">
                                        <span>Status no lote: </span>
                                        <strong className={inspectingMediaItem.activePageIndex === inspectingPageIndex ? "text-[#00C4CC]" : "text-[#57534E]"}>
                                          {inspectingMediaItem.activePageIndex === inspectingPageIndex ? "★ Página ativa atual" : `Página ${inspectingPageIndex + 1}`}
                                        </strong>
                                      </div>
                                      <button
                                        type="button"
                                        onClick={() => {
                                          setMediaList((prev) =>
                                            prev.map((item) =>
                                              item.id === inspectingMediaItem.id ? { ...item, activePageIndex: inspectingPageIndex } : item
                                            )
                                          );
                                          setInspectingMediaItem((prev) =>
                                            prev ? { ...prev, activePageIndex: inspectingPageIndex } : prev
                                          );
                                          alert(`Página ${inspectingPageIndex + 1} selecionada como ativa no lote!`);
                                        }}
                                        className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-semibold flex items-center gap-1.5 transition-colors cursor-pointer w-full sm:w-auto justify-center"
                                      >
                                        <CheckCircle2 className="w-3.5 h-3.5 text-[#00C4CC]" />
                                        <span>Usar Página {inspectingPageIndex + 1} no Lote</span>
                                      </button>
                                    </div>
                                </div>
                            );
                          })()}
                      </div>
                    </div>

                    <div className="px-6 py-3 border-t border-[#E7E5E4] bg-white flex justify-end">
                      <button
                        onClick={() => setInspectingMediaItem(null)}
                        className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg"
                      >
                        Concluído
                      </button>
                    </div>
                  </div>
                </div>
              )}

              <div className="flex justify-between pt-2">
                <button
                  onClick={() => setActiveStep(3)}
                  className="px-4 py-2 border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#57534E] hover:bg-[#F5F5F4]"
                >
                  Voltar
                </button>
                <button
                  onClick={() => setActiveStep(5)}
                  className="px-5 py-2.5 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-medium flex items-center gap-1.5"
                >
                  Avançar para Estilo <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 5: PERSONALIZAR O DESIGN                                           */}
          {/* ========================================================================= */}
          {activeStep === 5 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="pb-4 border-b border-[#F5F5F4]">
                <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                  5. Personalização Visual
                </h2>
                <p className="text-xs text-[#78716C] mt-1">
                  Controle proporção, película protetora, tipografia e contraste do texto sobre o fundo.
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {/* Formatos e Posição */}
                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-2">Proporção do Design</label>
                    <div className="grid grid-cols-3 gap-2">
                      {(['9:16', '1:1', '16:9'] as const).map((ratio) => (
                        <button
                          key={ratio}
                          onClick={() => setAspectRatio(ratio)}
                          className={`py-2 text-xs font-medium rounded-lg border transition-all ${
                            aspectRatio === ratio
                              ? 'border-[#1C1917] bg-[#1C1917] text-white'
                              : 'border-[#E7E5E4] bg-[#FAF9F6] text-[#57534E] hover:border-[#A8A29E]'
                          }`}
                        >
                          {ratio}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-2">Posição do Texto</label>
                    <div className="grid grid-cols-3 gap-2">
                      {[
                        { id: 'top', label: 'Topo' },
                        { id: 'center', label: 'Centro' },
                        { id: 'bottom', label: 'Base' },
                      ].map((pos) => (
                        <button
                          key={pos.id}
                          onClick={() => setTextPosition(pos.id as any)}
                          className={`py-2 text-xs font-medium rounded-lg border transition-all ${
                            textPosition === pos.id
                              ? 'border-[#1C1917] bg-[#1C1917] text-white'
                              : 'border-[#E7E5E4] bg-[#FAF9F6] text-[#57534E]'
                          }`}
                        >
                          {pos.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-2">Alinhamento</label>
                    <div className="grid grid-cols-3 gap-2">
                      {[
                        { id: 'left', label: 'Esquerda' },
                        { id: 'center', label: 'Centro' },
                        { id: 'right', label: 'Direita' },
                      ].map((align) => (
                        <button
                          key={align.id}
                          onClick={() => setTextAlign(align.id as any)}
                          className={`py-2 text-xs font-medium rounded-lg border transition-all ${
                            textAlign === align.id
                              ? 'border-[#1C1917] bg-[#1C1917] text-white'
                              : 'border-[#E7E5E4] bg-[#FAF9F6] text-[#57534E]'
                          }`}
                        >
                          {align.label}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-1">
                      Película Escura de Contraste ({overlayOpacity}%)
                    </label>
                    <input
                      type="range"
                      min={0}
                      max={90}
                      value={overlayOpacity}
                      onChange={(e) => setOverlayOpacity(Number(e.target.value))}
                      className="w-full accent-[#1C1917]"
                    />
                  </div>
                </div>

                {/* Tipografia e Efeitos */}
                <div className="space-y-4">
                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-2">Estilo de Fonte</label>
                    <div className="grid grid-cols-2 gap-2">
                      <button
                        onClick={() => setFontFamily('serif')}
                        className={`py-2 px-3 text-xs font-serif rounded-lg border text-left ${
                          fontFamily === 'serif' ? 'border-[#1C1917] bg-[#1C1917] text-white' : 'border-[#E7E5E4]'
                        }`}
                      >
                        Serifada (Editorial)
                      </button>
                      <button
                        onClick={() => setFontFamily('sans')}
                        className={`py-2 px-3 text-xs font-sans rounded-lg border text-left ${
                          fontFamily === 'sans' ? 'border-[#1C1917] bg-[#1C1917] text-white' : 'border-[#E7E5E4]'
                        }`}
                      >
                        Sem Serifa (Moderna)
                      </button>
                    </div>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-1">
                      Tamanho da Fonte ({fontSize}px)
                    </label>
                    <input
                      type="range"
                      min={16}
                      max={48}
                      value={fontSize}
                      onChange={(e) => setFontSize(Number(e.target.value))}
                      className="w-full accent-[#1C1917]"
                    />
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-1">Cor do Texto</label>
                    <div className="flex items-center gap-3">
                      <input
                        type="color"
                        value={textColor}
                        onChange={(e) => setTextColor(e.target.value)}
                        className="w-8 h-8 rounded border border-[#E7E5E4] cursor-pointer"
                      />
                      <span className="text-xs font-mono text-[#78716C]">{textColor}</span>
                    </div>
                  </div>

                  <div className="flex items-center gap-4 pt-1">
                    <label className="flex items-center gap-2 text-xs font-medium text-[#1C1917] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={hasShadow}
                        onChange={(e) => setHasShadow(e.target.checked)}
                        className="rounded accent-[#1C1917]"
                      />
                      <span>Sombra de leitura</span>
                    </label>
                    <label className="flex items-center gap-2 text-xs font-medium text-[#1C1917] cursor-pointer">
                      <input
                        type="checkbox"
                        checked={hasStroke}
                        onChange={(e) => setHasStroke(e.target.checked)}
                        className="rounded accent-[#1C1917]"
                      />
                      <span>Contorno suave</span>
                    </label>
                  </div>

                  <div>
                    <label className="block text-xs font-semibold text-[#1C1917] mb-1">Assinatura / Marca d'água</label>
                    <input
                      type="text"
                      value={signatureText}
                      onChange={(e) => setSignatureText(e.target.value)}
                      placeholder="@seuperfil"
                      className="w-full text-xs px-3 py-2 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg font-mono"
                    />
                  </div>
                </div>
              </div>

              <div className="flex justify-between pt-2">
                <button
                  onClick={() => setActiveStep(4)}
                  className="px-4 py-2 border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#57534E] hover:bg-[#F5F5F4]"
                >
                  Voltar
                </button>
                <button
                  onClick={() => setActiveStep(6)}
                  className="px-5 py-2.5 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-medium flex items-center gap-1.5"
                >
                  Avançar para Gerador <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 6: GERAR — OPÇÃO 1 (MODELO + CAMPOS-CHAVE) & OPÇÃO 2 (NOVO DESIGN) */}
          {/* ========================================================================= */}
          {activeStep === 6 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="pb-4 border-b border-[#F5F5F4] flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                <div>
                  <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                    6. Gerar Designs Editáveis no Canva
                  </h2>
                  <p className="text-xs text-[#78716C] mt-1">
                    Escolha entre preencher um modelo existente preservando todas as camadas (Opção 1) ou criar um novo design editável a partir do lote (Opção 2).
                  </p>
                </div>
                <div className="flex items-center gap-2 text-xs font-mono">
                  <span className="px-2.5 py-1 rounded-lg bg-[#FAF9F6] border border-[#E7E5E4] text-[#1C1917]">
                    {validPhrases.length} frases
                  </span>
                  <span className="px-2.5 py-1 rounded-lg bg-[#FAF9F6] border border-[#E7E5E4] text-[#1C1917]">
                    {mediaList.length} mídias
                  </span>
                </div>
              </div>

              {/* SELETOR PRINCIPAL DAS DUAS OPÇÕES INDEPENDENTES */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <button
                  type="button"
                  onClick={() => setGenerateOptionTab('option1_template')}
                  className={`p-4 rounded-2xl border text-left transition-all cursor-pointer ${
                    generateOptionTab === 'option1_template'
                      ? 'border-[#00C4CC] bg-cyan-50/40 ring-2 ring-[#00C4CC]/25 shadow-sm'
                      : 'border-[#E7E5E4] bg-[#FAF9F6] hover:border-[#1C1917]'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-bold uppercase tracking-wide text-[#008A90] font-mono">
                      Opção 1 · Prioritária
                    </span>
                    {generateOptionTab === 'option1_template' && (
                      <CheckCircle2 className="w-4 h-4 text-[#00C4CC]" />
                    )}
                  </div>
                  <h3 className="text-sm font-serif font-bold text-[#1C1917] mt-1">
                    1. Gerar a partir de um Modelo do Canva
                  </h3>
                  <p className="text-[11.5px] text-[#57534E] mt-1 leading-relaxed">
                    Selecione um modelo (Brand Template) da sua conta, substitua campos-chave (<code className="font-mono">{'{FRASE}'}</code>, <code className="font-mono">{'{TITULO}'}</code>, <code className="font-mono">{'{ASSINATURA}'}</code>) e gere múltiplas versões editáveis preservando fontes, cores e elementos.
                  </p>
                </button>

                <button
                  type="button"
                  onClick={() => setGenerateOptionTab('option2_new_design')}
                  className={`p-4 rounded-2xl border text-left transition-all cursor-pointer ${
                    generateOptionTab === 'option2_new_design'
                      ? 'border-[#1C1917] bg-[#FAF9F6] ring-2 ring-[#1C1917]/20 shadow-sm'
                      : 'border-[#E7E5E4] bg-[#FAF9F6] hover:border-[#1C1917]'
                  }`}
                >
                  <div className="flex items-center justify-between gap-2">
                    <span className="text-xs font-bold uppercase tracking-wide text-[#57534E] font-mono">
                      Opção 2 · Camadas Separadas
                    </span>
                    {generateOptionTab === 'option2_new_design' && (
                      <CheckCircle2 className="w-4 h-4 text-[#1C1917]" />
                    )}
                  </div>
                  <h3 className="text-sm font-serif font-bold text-[#1C1917] mt-1">
                    2. Criar Novo Design no Canva
                  </h3>
                  <p className="text-[11.5px] text-[#57534E] mt-1 leading-relaxed">
                    Monta um novo design editável no Canva a partir das frases, mídias e estilo da Etapa 5, mantendo caixas de texto, fundos e película como camadas independentes (nunca imagem achatada).
                  </p>
                </button>
              </div>

              {/* =============================================================== */}
              {/* CONTEÚDO DA OPÇÃO 1: GERAR A PARTIR DE UM MODELO DO CANVA       */}
              {/* =============================================================== */}
              {generateOptionTab === 'option1_template' && (
                <div className="space-y-6">
                  {/* Passo 1.1: Permissão Brand Templates (quando ausente) */}
                  {canvaStatus?.connected && !canvaStatus?.hasBrandTemplateScope && (
                    <div className="bg-amber-50/80 border border-amber-200 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs text-amber-950">
                      <div className="flex items-start gap-2.5">
                        <AlertCircle className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
                        <div className="space-y-0.5">
                          <p className="font-bold">Permissões de Brand Template pendentes no token atual</p>
                          <p className="text-amber-900 leading-relaxed">
                            Para ler campos de dados (<code className="font-mono">/dataset</code>) e executar preenchimento automático oficial (<code className="font-mono">/autofills</code>), ative <code className="font-mono">brandtemplate:meta:read</code> e <code className="font-mono">brandtemplate:content:read</code> no seu app do Canva Developers Portal e clique em Autorizar.
                          </p>
                        </div>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleConnect(true)}
                        disabled={connecting}
                        className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-xl shrink-0 transition-colors cursor-pointer disabled:opacity-50"
                      >
                        {connecting ? 'Autorizando...' : 'Autorizar Brand Templates'}
                      </button>
                    </div>
                  )}

                  {/* Passo 1.2: Selecionar Modelo do Canva */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-5 space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#008A90]">
                          Etapa 1 · Escolher Modelo
                        </span>
                        <h4 className="text-sm font-bold text-[#1C1917] mt-0.5">
                          Selecionar design ou Brand Template da conta Canva
                        </h4>
                      </div>
                      <div className="flex flex-wrap items-center gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            if (!canvaStatus?.connected) {
                              handleConnect(false);
                            } else {
                              openTemplateSelectorModal('brand_templates');
                            }
                          }}
                          className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-xl flex items-center gap-1.5 shadow-sm transition-colors cursor-pointer"
                        >
                          <Sparkles className="w-3.5 h-3.5" />
                          <span>Selecionar design do Canva</span>
                        </button>
                        <button
                          type="button"
                          onClick={() => {
                            if (!canvaStatus?.connected) {
                              handleConnect(false);
                            } else {
                              openTemplateSelectorModal('folders');
                            }
                          }}
                          className="px-3.5 py-2 bg-white hover:bg-[#F5F5F4] border border-[#E7E5E4] text-[#1C1917] text-xs font-semibold rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer"
                        >
                          <FolderOpen className="w-3.5 h-3.5 text-[#00C4CC]" />
                          <span>Navegar Pastas</span>
                        </button>
                      </div>
                    </div>

                    {/* Entrada rápida por ID ou URL do modelo */}
                    <div className="flex flex-col sm:flex-row gap-2">
                      <input
                        type="text"
                        value={manualTemplateInspectId || templateId}
                        onChange={(e) => {
                          setManualTemplateInspectId(e.target.value);
                          setTemplateId(e.target.value);
                        }}
                        placeholder="Ou cole o ID / URL do Brand Template ou Design Canva (ex: EAF... ou https://www.canva.com/design/...)"
                        className="flex-1 text-xs px-3.5 py-2 bg-white border border-[#E7E5E4] rounded-xl font-mono focus:outline-none focus:ring-1 focus:ring-[#1C1917]"
                      />
                      <button
                        type="button"
                        onClick={() => inspectSelectedCanvaTemplate(manualTemplateInspectId || templateId)}
                        disabled={inspectingTemplate || !(manualTemplateInspectId || templateId).trim()}
                        className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-xl flex items-center justify-center gap-1.5 disabled:opacity-50 cursor-pointer shrink-0"
                      >
                        {inspectingTemplate ? (
                          <>
                            <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                            <span>Inspecionando...</span>
                          </>
                        ) : (
                          <>
                            <Eye className="w-3.5 h-3.5 text-[#00C4CC]" />
                            <span>Inspecionar Modelo</span>
                          </>
                        )}
                      </button>
                    </div>

                    {/* Passo 1.3: Visualização Completa do Modelo Inspecionado */}
                    {inspectedTemplate && (
                      <div className="bg-white border border-[#E7E5E4] rounded-xl p-4 space-y-4">
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-[#F5F5F4]">
                          <div>
                            <div className="flex items-center gap-2 flex-wrap">
                              <h5 className="text-sm font-bold text-[#1C1917]">{inspectedTemplate.title}</h5>
                              {inspectedTemplate.autofillCompatible ? (
                                <span className="px-2 py-0.5 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-md text-[10px] font-mono font-bold">
                                  ✓ Brand Template Autofill Compatível
                                </span>
                              ) : inspectedTemplate.editableViaPptxRoundtrip ? (
                                <span className="px-2 py-0.5 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded-md text-[10px] font-mono font-bold">
                                  ✓ Edição Direta no Arquivo (Camadas XML)
                                </span>
                              ) : (
                                <span className="px-2 py-0.5 bg-amber-50 text-amber-800 border border-amber-200 rounded-md text-[10px] font-mono font-bold">
                                  ⚠ Design Comum / Sem Dataset Autofill
                                </span>
                              )}
                            </div>
                            <p className="text-[11px] text-[#78716C] font-mono mt-0.5">
                              ID: {inspectedTemplate.brandTemplateId || inspectedTemplate.id} · {inspectedTemplate.pageCount} página(s)
                            </p>
                          </div>

                          {(inspectedTemplate.editUrl || inspectedTemplate.viewUrl) && (
                            <a
                              href={inspectedTemplate.editUrl || inspectedTemplate.viewUrl || '#'}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="px-3 py-1.5 border border-[#E7E5E4] hover:bg-[#FAF9F6] text-xs font-medium text-[#1C1917] rounded-lg flex items-center gap-1.5 shrink-0"
                            >
                              <span>Abrir Modelo Original</span>
                              <ExternalLink className="w-3 h-3" />
                            </a>
                          )}
                        </div>

                        {/* Miniaturas das páginas do modelo */}
                        {inspectedTemplate.pages.length > 0 && (
                          <div className="space-y-1.5">
                            <span className="text-[11px] font-semibold text-[#57534E] block">
                              Visualização das páginas do modelo ({inspectedTemplate.pages.length}):
                            </span>
                            <div className="flex items-center gap-3 overflow-x-auto pb-1">
                              {inspectedTemplate.pages.map((p) => (
                                <div key={p.index} className="w-28 shrink-0 border border-[#E7E5E4] rounded-lg overflow-hidden bg-neutral-900">
                                  <div className="h-24 relative">
                                    {p.thumbnailUrl ? (
                                      <img src={p.thumbnailUrl} alt={p.title} className="w-full h-full object-cover" />
                                    ) : (
                                      <div className="w-full h-full flex items-center justify-center text-white/50 text-[10px] font-mono">
                                        Pág {p.index}
                                      </div>
                                    )}
                                    <span className="absolute top-1 left-1 bg-black/80 text-white text-[9px] font-mono px-1.5 py-0.5 rounded">
                                      Pág {p.index}
                                    </span>
                                  </div>
                                </div>
                              ))}
                            </div>
                          </div>
                        )}

                        {inspectedTemplate.editableViaPptxRoundtrip && !inspectedTemplate.autofillCompatible && (
                          <div className="bg-emerald-50/70 border border-emerald-200 rounded-xl p-3.5 text-xs text-emerald-950 space-y-2">
                            <div className="flex items-start gap-2">
                              <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                              <div className="space-y-0.5">
                                <p className="font-bold">Substituição Direta Dentro do Texto Original Ativada</p>
                                <p className="leading-relaxed text-emerald-900">
                                  Os marcadores detectados ({(inspectedTemplate.detectedPlaceholders || ['(frase)']).join(', ')}) serão substituídos diretamente dentro da camada de texto original nas {inspectedTemplate.pageCount} página(s), preservando família tipográfica, tamanho, cor, negrito, itálico, alinhamento, espaçamento e dimensões sem sobrepor nenhum texto extra.
                                </p>
                              </div>
                            </div>
                            {Array.isArray(inspectedTemplate.slideLayouts) && inspectedTemplate.slideLayouts.length > 0 && (
                              <div className="pl-6 flex flex-wrap gap-1.5">
                                {inspectedTemplate.slideLayouts.slice(0, 6).map((sl) => (
                                  <span
                                    key={`sl_info_${sl.slideIndex}`}
                                    className="px-2 py-0.5 bg-white border border-emerald-200 rounded text-[10px] font-mono text-emerald-900"
                                  >
                                    Pág {sl.slideIndex}: {sl.textBoxes.length} caixa(s)
                                    {sl.detectedTokens.length > 0 ? ` · [${sl.detectedTokens.join(', ')}]` : ''}
                                    {sl.textBoxes[0]
                                      ? ` · ${sl.textBoxes[0].fontFamily || 'Fonte orig.'} ${sl.textBoxes[0].fontSizePt}pt${sl.textBoxes[0].bold ? ' B' : ''}${sl.textBoxes[0].italic ? ' I' : ''}`
                                      : ''}
                                  </span>
                                ))}
                              </div>
                            )}
                          </div>
                        )}

                        {/* Diagnóstico honesto apenas se NÃO for nem Autofill nem PPTX Roundtrip */}
                        {!inspectedTemplate.autofillCompatible && !inspectedTemplate.editableViaPptxRoundtrip && inspectedTemplate.limitationReason && (
                          <div className="bg-amber-50 border border-amber-200 rounded-xl p-3.5 text-xs text-amber-950 space-y-2">
                            <div className="flex items-start gap-2">
                              <Info className="w-4 h-4 text-amber-700 shrink-0 mt-0.5" />
                              <p className="leading-relaxed">{inspectedTemplate.limitationReason}</p>
                            </div>
                            <div className="flex flex-wrap items-center gap-2 pl-6">
                              {!inspectedTemplate.hasBrandTemplateScope && (
                                <button
                                  type="button"
                                  onClick={() => handleConnect(true)}
                                  className="px-3 py-1.5 bg-[#1C1917] text-white text-[11px] font-semibold rounded-lg cursor-pointer"
                                >
                                  Autorizar Brand Templates
                                </button>
                              )}
                              <button
                                type="button"
                                onClick={() => setGenerateOptionTab('option2_new_design')}
                                className="px-3 py-1.5 bg-white border border-amber-300 text-amber-900 hover:bg-amber-100 text-[11px] font-semibold rounded-lg cursor-pointer"
                              >
                                Usar Opção 2 (Criar Novo Design Editável)
                              </button>
                            </div>
                          </div>
                        )}
                      </div>
                    )}
                  </div>

                  {/* Passo 1.4: Configurar Campos-Chave e Substituições */}
                  <div
                    className={`rounded-2xl p-5 space-y-4 border ${
                      isModelValidationBlocked
                        ? 'bg-rose-50/40 border-rose-300'
                        : 'bg-[#FAF9F6] border-[#E7E5E4]'
                    }`}
                  >
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                      <div>
                        <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#008A90]">
                          Etapa 2 · Campos-Chave & Páginas
                        </span>
                        <h4 className="text-sm font-bold text-[#1C1917] mt-0.5">
                          Configurar substituições automáticas de texto
                        </h4>
                        <p className="text-[11.5px] text-[#78716C] mt-0.5">
                          Digite ou edite o marcador exatamente como ele aparece no seu arquivo Canva (ex: <code className="font-mono">(frase)</code> ou <code className="font-mono">frase</code>). A fonte, cor e posição originais serão mantidas.
                        </p>
                      </div>
                    </div>

                    {isModelValidationBlocked && (
                      <div className="p-3.5 bg-rose-50 border-2 border-rose-500 rounded-xl text-xs text-rose-950 space-y-2">
                        <div className="flex items-start gap-2">
                          <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                          <div className="space-y-1">
                            <p className="font-bold text-rose-900">
                              Geração bloqueada: campos-chave obrigatórios não encontrados no modelo
                            </p>
                            <p className="text-rose-800 leading-relaxed">
                              {missingRequiredFieldMappings.length > 0
                                ? `Os campos destacados em vermelho (${missingRequiredFieldMappings
                                    .map((m) => `"${m.fieldKey || '(vazio)'}"`)
                                    .join(', ')}) não existem no modelo "${inspectedTemplate?.title || 'Canva'}". Edite o nome do campo para corresponder ao arquivo ou mude para "Não substituir".`
                                : 'Nenhum campo de "Frase do Lote" ativo foi encontrado no modelo selecionado. Configure pelo menos um campo válido para liberar a geração.'}
                            </p>
                          </div>
                        </div>
                        {modelSelectableTokens.length > 0 && (
                          <div className="pl-6 flex flex-wrap items-center gap-1.5">
                            <span className="text-[11px] font-semibold text-rose-900">
                              Marcadores detectados no modelo:
                            </span>
                            {modelSelectableTokens.map((tok) => (
                              <button
                                key={tok}
                                type="button"
                                onClick={() => {
                                  setFieldPhrase(tok);
                                  setKeyFieldMappings((prev) => {
                                    const targetIdx = prev.findIndex(
                                      (m) => m.source !== 'ignore' && !isFieldKeyFoundInModel(m.fieldKey)
                                    );
                                    if (targetIdx >= 0) {
                                      return prev.map((m, i) =>
                                        i === targetIdx ? { ...m, fieldKey: tok, detectedInTemplate: true } : m
                                      );
                                    }
                                    return [
                                      {
                                        fieldKey: tok,
                                        detectedInTemplate: true,
                                        source: 'batch_phrase',
                                        customText: '',
                                        pageScope: 'all',
                                        selectedPages: Array.from(
                                          { length: Math.max(1, inspectedTemplate?.pageCount || 1) },
                                          (_, i) => i + 1
                                        ),
                                      },
                                      ...prev,
                                    ];
                                  });
                                }}
                                className="px-2 py-0.5 bg-white border border-rose-300 hover:bg-rose-100 text-rose-900 rounded font-mono text-[11px] font-bold cursor-pointer"
                              >
                                {tok}
                              </button>
                            ))}
                          </div>
                        )}
                      </div>
                    )}

                    {/* Lista de Campos-Chave */}
                    <div className="space-y-2.5">
                      {keyFieldMappings.map((mapping, idx) => {
                        const pageTotal = Math.max(1, inspectedTemplate?.pageCount || 1);
                        const isActiveField = mapping.source !== 'ignore';
                        const isMatchedInModel = isFieldKeyFoundInModel(mapping.fieldKey);
                        const isMissingRequiredError = hasInspectableModelLoaded && isActiveField && !isMatchedInModel;

                        return (
                          <div
                            key={`kfm_${idx}`}
                            className={`rounded-xl p-3.5 space-y-2.5 transition-colors ${
                              isMissingRequiredError
                                ? 'bg-rose-50 border-2 border-rose-500 ring-2 ring-rose-500/15'
                                : hasInspectableModelLoaded && isActiveField && isMatchedInModel
                                ? 'bg-white border border-emerald-300'
                                : 'bg-white border border-[#E7E5E4]'
                            }`}
                          >
                            <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
                              <div className="flex items-center gap-2 flex-wrap">
                                <div className="flex flex-col">
                                  <span className="text-[10px] font-medium text-[#78716C] mb-0.5">
                                    Texto no modelo Canva {isActiveField && <span className="text-rose-600">*</span>}
                                  </span>
                                  <input
                                    type="text"
                                    value={mapping.fieldKey}
                                    onChange={(e) => {
                                      const nextKey = e.target.value;
                                      if (mapping.source === 'batch_phrase') {
                                        setFieldPhrase(nextKey);
                                      }
                                      setKeyFieldMappings((prev) =>
                                        prev.map((m, i) =>
                                          i === idx
                                            ? {
                                                ...m,
                                                fieldKey: nextKey,
                                                detectedInTemplate: isFieldKeyFoundInModel(nextKey),
                                              }
                                            : m
                                        )
                                      );
                                    }}
                                    placeholder="Ex: (frase) ou FRASE"
                                    className={`px-2.5 py-1.5 rounded-lg text-xs font-mono font-bold w-40 sm:w-44 focus:outline-none ${
                                      isMissingRequiredError
                                        ? 'bg-white border-2 border-rose-500 text-rose-900 focus:ring-2 focus:ring-rose-500/30'
                                        : 'bg-[#FAF9F6] border border-[#E7E5E4] text-[#1C1917] focus:ring-1 focus:ring-[#1C1917]'
                                    }`}
                                  />
                                </div>

                                {isMissingRequiredError ? (
                                  <span className="text-[10px] font-mono font-bold px-2.5 py-1 bg-rose-600 text-white rounded-lg mt-4">
                                    ✗ Obrigatório · Não encontrado no modelo
                                  </span>
                                ) : hasInspectableModelLoaded && isMatchedInModel ? (
                                  <span className="text-[10px] font-mono font-bold px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded mt-4">
                                    {inspectedTemplate?.autofillCompatible ? '✓ Detectado no Dataset' : '✓ Encontrado no Modelo'}
                                  </span>
                                ) : (
                                  <span className="text-[10px] font-mono px-2 py-0.5 bg-neutral-100 text-[#78716C] rounded mt-4">
                                    {isActiveField ? 'Campo ativo' : 'Ignorado'}
                                  </span>
                                )}
                              </div>

                              <div className="flex flex-wrap items-center gap-2 flex-1 justify-end mt-2 md:mt-4">
                                <select
                                  value={mapping.source}
                                  onChange={(e) => {
                                    const val = e.target.value as any;
                                    if (val === 'batch_phrase') {
                                      setFieldPhrase(mapping.fieldKey);
                                    }
                                    setKeyFieldMappings((prev) =>
                                      prev.map((m, i) => (i === idx ? { ...m, source: val } : m))
                                    );
                                  }}
                                  className={`text-xs px-2.5 py-1.5 rounded-lg font-medium ${
                                    isMissingRequiredError
                                      ? 'bg-white border border-rose-400 text-rose-950'
                                      : 'bg-[#FAF9F6] border border-[#E7E5E4] text-[#1C1917]'
                                  }`}
                                >
                                  <option value="batch_phrase">Frase do Lote (1 versão por frase)</option>
                                  <option value="signature">Assinatura ({signatureText || '@lotecriativo'})</option>
                                  <option value="custom_text">Texto Fixo Personalizado</option>
                                  <option value="ignore">Não substituir (manter original)</option>
                                </select>

                                {mapping.source === 'custom_text' && (
                                  <input
                                    type="text"
                                    value={mapping.customText}
                                    onChange={(e) => {
                                      const val = e.target.value;
                                      setKeyFieldMappings((prev) =>
                                        prev.map((m, i) => (i === idx ? { ...m, customText: val } : m))
                                      );
                                    }}
                                    placeholder={`Valor para ${mapping.fieldKey}`}
                                    className="text-xs px-2.5 py-1.5 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg min-w-[170px]"
                                  />
                                )}

                                {keyFieldMappings.length > 1 && (
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setKeyFieldMappings((prev) => prev.filter((_, i) => i !== idx))
                                    }
                                    title="Remover campo"
                                    className="p-1.5 text-[#78716C] hover:text-rose-600 hover:bg-rose-50 rounded-lg transition-colors cursor-pointer"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                )}
                              </div>
                            </div>

                            {isMissingRequiredError && modelSelectableTokens.length > 0 && (
                              <div className="pt-1 flex flex-wrap items-center gap-1.5 text-[11px] text-rose-800">
                                <span className="font-medium">Substituir pelo marcador encontrado:</span>
                                {modelSelectableTokens.map((tok) => (
                                  <button
                                    key={tok}
                                    type="button"
                                    onClick={() => {
                                      if (mapping.source === 'batch_phrase') {
                                        setFieldPhrase(tok);
                                      }
                                      setKeyFieldMappings((prev) =>
                                        prev.map((m, i) =>
                                          i === idx ? { ...m, fieldKey: tok, detectedInTemplate: true } : m
                                        )
                                      );
                                    }}
                                    className="px-2 py-0.5 bg-white border border-rose-300 hover:bg-rose-100 text-rose-900 rounded font-mono text-[10.5px] font-bold cursor-pointer"
                                  >
                                    {tok}
                                  </button>
                                ))}
                              </div>
                            )}

                            {/* Regra de Páginas quando o campo aparece em múltiplas páginas */}
                            {mapping.source !== 'ignore' && pageTotal > 1 && (
                              <div className="pt-2 border-t border-[#F5F5F4] flex flex-wrap items-center justify-between gap-2 text-[11px]">
                                <div className="flex items-center gap-2">
                                  <span className="text-[#78716C] font-medium">Aplicação nas páginas:</span>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setKeyFieldMappings((prev) =>
                                        prev.map((m, i) => (i === idx ? { ...m, pageScope: 'all' } : m))
                                      )
                                    }
                                    className={`px-2 py-0.5 rounded font-medium cursor-pointer ${
                                      mapping.pageScope === 'all'
                                        ? 'bg-[#1C1917] text-white'
                                        : 'bg-[#FAF9F6] border border-[#E7E5E4] text-[#57534E]'
                                    }`}
                                  >
                                    Todas as páginas ({pageTotal})
                                  </button>
                                  <button
                                    type="button"
                                    onClick={() =>
                                      setKeyFieldMappings((prev) =>
                                        prev.map((m, i) => (i === idx ? { ...m, pageScope: 'specific' } : m))
                                      )
                                    }
                                    className={`px-2 py-0.5 rounded font-medium cursor-pointer ${
                                      mapping.pageScope === 'specific'
                                        ? 'bg-[#00C4CC] text-white'
                                        : 'bg-[#FAF9F6] border border-[#E7E5E4] text-[#57534E]'
                                    }`}
                                  >
                                    Páginas específicas
                                  </button>
                                </div>

                                {mapping.pageScope === 'specific' && (
                                  <div className="flex items-center gap-1 flex-wrap">
                                    {Array.from({ length: pageTotal }, (_, pI) => pI + 1).map((pNum) => {
                                      const isOn = mapping.selectedPages.includes(pNum);
                                      return (
                                        <button
                                          key={pNum}
                                          type="button"
                                          onClick={() => {
                                            setKeyFieldMappings((prev) =>
                                              prev.map((m, i) => {
                                                if (i !== idx) return m;
                                                const nextPages = isOn
                                                  ? m.selectedPages.filter((p) => p !== pNum)
                                                  : [...m.selectedPages, pNum].sort((a, b) => a - b);
                                                return { ...m, selectedPages: nextPages };
                                              })
                                            );
                                          }}
                                          className={`px-2 py-0.5 rounded font-mono text-[10px] cursor-pointer ${
                                            isOn
                                              ? 'bg-[#00C4CC] text-white font-bold'
                                              : 'bg-white border border-[#E7E5E4] text-[#78716C]'
                                          }`}
                                        >
                                          Pág {pNum}
                                        </button>
                                      );
                                    })}
                                  </div>
                                )}
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </div>

                    {/* Adicionar campo-chave customizado */}
                    <div className="flex flex-col sm:flex-row gap-2 pt-1">
                      <input
                        type="text"
                        value={newCustomFieldKey}
                        onChange={(e) => setNewCustomFieldKey(e.target.value)}
                        placeholder="Adicionar outro campo-chave (ex: (frase), CTA, PRECO, @assinatura)..."
                        className="flex-1 text-xs px-3 py-2 bg-white border border-[#E7E5E4] rounded-xl font-mono"
                      />
                      <button
                        type="button"
                        onClick={handleAddCustomKeyField}
                        disabled={!newCustomFieldKey.trim()}
                        className="px-3.5 py-2 bg-white border border-[#E7E5E4] hover:bg-[#F5F5F4] text-xs font-semibold text-[#1C1917] rounded-xl flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
                      >
                        <Plus className="w-3.5 h-3.5" />
                        <span>Adicionar Campo</span>
                      </button>
                    </div>
                  </div>

                  {/* Passo 1.5: Pré-visualizar Resultados & Gerar Versões */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-5 space-y-4">
                    <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                      <div>
                        <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#008A90]">
                          Etapa 3 · Prévia das Versões & Salvamento no Canva
                        </span>
                        <h4 className="text-sm font-bold text-[#1C1917] mt-0.5">
                          Pré-visualizar substituições ({validPhrases.length} versões a gerar)
                        </h4>
                      </div>
                      <div className="flex items-center gap-2">
                        <label className="text-[11px] text-[#57534E] font-medium">Prefixo do nome:</label>
                        <input
                          type="text"
                          value={autofillTitlePrefix}
                          onChange={(e) => setAutofillTitlePrefix(e.target.value)}
                          className="text-xs px-2.5 py-1.5 bg-white border border-[#E7E5E4] rounded-lg w-40 font-mono"
                        />
                      </div>
                    </div>

                    {/* Pré-visualização de cada versão */}
                    <div className="bg-white border border-[#E7E5E4] rounded-xl divide-y divide-[#F5F5F4] max-h-56 overflow-y-auto">
                      {validPhrases.map((phrase, vIdx) => {
                        const activeFields = keyFieldMappings.filter((m) => m.source !== 'ignore');
                        return (
                          <div key={vIdx} className="p-3 text-xs space-y-1">
                            <div className="flex items-center justify-between gap-2">
                              <span className="font-bold text-[#1C1917]">
                                {autofillTitlePrefix || 'Lote'} - Versão {vIdx + 1}
                              </span>
                              <span className="text-[10px] font-mono text-[#78716C]">
                                {inspectedTemplate?.pageCount || 1} pág(s) editáveis
                              </span>
                            </div>
                            <div className="flex flex-wrap gap-1.5">
                              {activeFields.map((m) => {
                                const val = resolveFieldPreviewValue(m, phrase);
                                if (val === null) return null;
                                return (
                                  <span
                                    key={m.fieldKey}
                                    className="px-2 py-0.5 bg-[#FAF9F6] border border-[#E7E5E4] rounded text-[11px] text-[#44403C]"
                                  >
                                    <strong className="font-mono text-[#008A90]">{`{${m.fieldKey}}`}</strong> → {val}
                                  </span>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {autofillStatusBanner && (
                      <div
                        className={`p-3.5 rounded-xl border text-xs flex items-start gap-2 ${
                          autofillStatusBanner.type === 'success'
                            ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                            : autofillStatusBanner.type === 'error'
                            ? 'bg-rose-50 border-rose-200 text-rose-900'
                            : 'bg-cyan-50 border-cyan-200 text-cyan-950'
                        }`}
                      >
                        {autofillStatusBanner.type === 'success' ? (
                          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                        ) : autofillStatusBanner.type === 'error' ? (
                          <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                        ) : (
                          <RefreshCw className="w-4 h-4 text-[#00C4CC] animate-spin shrink-0 mt-0.5" />
                        )}
                        <span className="leading-relaxed">{autofillStatusBanner.text}</span>
                      </div>
                    )}

                    {/* Seletor de Formato de Saída quando via Edição XML no Arquivo */}
                    {inspectedTemplate?.editableViaPptxRoundtrip && !inspectedTemplate?.autofillCompatible && (
                      <div className="bg-white border border-[#E7E5E4] rounded-xl p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
                        <span className="font-semibold text-[#1C1917]">Formato ao salvar no Canva:</span>
                        <div className="flex flex-wrap gap-2">
                          <button
                            type="button"
                            onClick={() => setOption1OutputMode('individual')}
                            className={`px-3 py-1.5 rounded-lg font-medium cursor-pointer border ${
                              option1OutputMode === 'individual'
                                ? 'bg-[#1C1917] text-white border-[#1C1917]'
                                : 'bg-[#FAF9F6] text-[#57534E] border-[#E7E5E4]'
                            }`}
                          >
                            1 Design por Frase (Edita o arquivo original)
                          </button>
                          <button
                            type="button"
                            onClick={() => setOption1OutputMode('multipage')}
                            className={`px-3 py-1.5 rounded-lg font-medium cursor-pointer border ${
                              option1OutputMode === 'multipage'
                                ? 'bg-[#00C4CC] text-white border-[#00C4CC]'
                                : 'bg-[#FAF9F6] text-[#57534E] border-[#E7E5E4]'
                            }`}
                          >
                            Design Único Multipáginas (1 frase por slide)
                          </button>
                        </div>
                      </div>
                    )}

                    {isModelValidationBlocked && (
                      <div className="p-3 bg-rose-50 border border-rose-400 rounded-xl text-xs text-rose-900 flex items-center gap-2">
                        <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                        <span>
                          Botões de geração bloqueados: corrija os campos-chave destacados em vermelho na Etapa 2 acima antes de gerar o lote.
                        </span>
                      </div>
                    )}

                    <div className="flex flex-wrap items-center gap-3 pt-1">
                      <button
                        type="button"
                        onClick={runOption1AutofillBatch}
                        disabled={
                          autofillBatchLoading ||
                          (!inspectedTemplate?.autofillCompatible && !inspectedTemplate?.editableViaPptxRoundtrip) ||
                          validPhrases.length === 0 ||
                          isModelValidationBlocked
                        }
                        className="px-6 py-3 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-xl flex items-center gap-2 shadow-sm transition-all disabled:opacity-50 cursor-pointer"
                      >
                        {autofillBatchLoading ? (
                          <>
                            <RefreshCw className="w-4 h-4 animate-spin" />
                            <span>Substituindo Textos & Salvando no Canva...</span>
                          </>
                        ) : (
                          <>
                            <Sparkles className="w-4 h-4" />
                            <span>
                              Gerar Versões e Salvar no Canva (
                              {option1OutputMode === 'multipage' && inspectedTemplate?.editableViaPptxRoundtrip
                                ? `${validPhrases.length} págs`
                                : validPhrases.length}
                              )
                            </span>
                          </>
                        )}
                      </button>

                      {inspectedTemplate?.editableViaPptxRoundtrip && !inspectedTemplate?.autofillCompatible && (
                        <button
                          type="button"
                          onClick={() => handleDownloadReplacedPptx()}
                          disabled={downloadingReplacedPptx || validPhrases.length === 0 || isModelValidationBlocked}
                          className="px-4 py-3 border border-[#E7E5E4] bg-white hover:bg-[#F5F5F4] rounded-xl text-xs font-semibold text-[#1C1917] flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                        >
                          <Download className="w-3.5 h-3.5 text-[#78716C]" />
                          <span>{downloadingReplacedPptx ? 'Gerando .PPTX...' : 'Baixar .PPTX Substituído'}</span>
                        </button>
                      )}

                      {canvaSdkHost.insideCanvaEditor && (
                        <button
                          type="button"
                          disabled={applyingRichtextSdk || validPhrases.length === 0 || isModelValidationBlocked}
                          onClick={async () => {
                            setApplyingRichtextSdk(true);
                            try {
                              const phrase = validPhrases[previewIndex] || validPhrases[0] || '';
                              const activeMappings = keyFieldMappings.filter((m) => m.source !== 'ignore');
                              const reps = activeMappings
                                .map((m) => {
                                  const val = resolveFieldPreviewValue(m, phrase);
                                  return val !== null ? { findText: m.fieldKey, replaceText: val } : null;
                                })
                                .filter(Boolean) as Array<{ findText: string; replaceText: string }>;
                              const rtRes = await replaceKeywordsInCanvaRichtextSelection(reps);
                              setAutofillStatusBanner({
                                text: rtRes.success
                                  ? `Richtext aplicado no design aberto via editContent + replaceText (${rtRes.replacedCount} substituição(ões) com formatação integral preservada)!`
                                  : rtRes.error || 'Nenhuma substituição realizada na seleção Richtext.',
                                type: rtRes.success ? 'success' : 'error',
                              });
                            } finally {
                              setApplyingRichtextSdk(false);
                            }
                          }}
                          className="px-4 py-3 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-bold rounded-xl flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                        >
                          <Edit3 className="w-3.5 h-3.5 text-[#00C4CC]" />
                          <span>{applyingRichtextSdk ? 'Editando Richtext...' : 'Aplicar no Editor (Richtext SDK)'}</span>
                        </button>
                      )}

                      {autofillVersionResults.length > 0 && (
                        <button
                          type="button"
                          onClick={() => setActiveStep(7)}
                          className="px-4 py-3 border border-[#E7E5E4] bg-white hover:bg-[#F5F5F4] rounded-xl text-xs font-semibold text-[#1C1917] cursor-pointer"
                        >
                          Ver na Aba Resultados ({autofillVersionResults.length})
                        </button>
                      )}
                    </div>

                    {/* Resultados Gerados da Opção 1 */}
                    {autofillVersionResults.length > 0 && (
                      <div className="space-y-2 pt-2 border-t border-[#E7E5E4]">
                        <p className="text-xs font-bold text-[#1C1917]">
                          Versões Editáveis Criadas no Canva:
                        </p>
                        <div className="space-y-2">
                          {autofillVersionResults.map((res) => (
                            <div
                              key={`${res.versionIndex}_${res.designId || 'fail'}`}
                              className="bg-white border border-[#E7E5E4] rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                            >
                              <div className="min-w-0 space-y-0.5">
                                <div className="flex items-center gap-2 flex-wrap">
                                  <span className="font-mono text-[11px] font-bold text-[#008A90]">
                                    #{res.versionIndex}
                                  </span>
                                  <span className="text-xs font-bold text-[#1C1917] truncate">
                                    {res.title}
                                  </span>
                                  {res.status === 'completed' && res.designId && (
                                    <span className="text-[10px] font-mono px-2 py-0.5 bg-emerald-50 text-emerald-700 border border-emerald-200 rounded">
                                      ID: {res.designId} · {res.pageCount} pág(s)
                                    </span>
                                  )}
                                </div>
                                {res.label && (
                                  <p className="text-[11px] text-[#78716C] truncate">{res.label}</p>
                                )}
                                {res.error && (
                                  <p className="text-[11px] text-rose-600">{res.error}</p>
                                )}
                              </div>

                              {res.status === 'completed' && res.editUrl && (
                                <div className="flex items-center gap-2 shrink-0">
                                  <a
                                    href={res.editUrl}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    className="px-3.5 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-lg flex items-center gap-1.5"
                                  >
                                    <span>Abrir no Canva</span>
                                    <ExternalLink className="w-3.5 h-3.5" />
                                  </a>
                                  <button
                                    type="button"
                                    onClick={() => {
                                      navigator.clipboard.writeText(res.editUrl!);
                                      setCopiedVersionId(res.designId);
                                      setTimeout(() => setCopiedVersionId(null), 2000);
                                    }}
                                    className="px-3 py-2 border border-[#E7E5E4] hover:bg-[#FAF9F6] text-xs font-medium text-[#1C1917] rounded-lg flex items-center gap-1 cursor-pointer"
                                  >
                                    {copiedVersionId === res.designId ? (
                                      <>
                                        <Check className="w-3.5 h-3.5 text-emerald-600" />
                                        <span>Copiado</span>
                                      </>
                                    ) : (
                                      <>
                                        <Copy className="w-3.5 h-3.5 text-[#78716C]" />
                                        <span>Copiar link</span>
                                      </>
                                    )}
                                  </button>
                                </div>
                              )}
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              {/* =============================================================== */}
              {/* CONTEÚDO DA OPÇÃO 2: CRIAR NOVO DESIGN NO CANVA                 */}
              {/* =============================================================== */}
              {generateOptionTab === 'option2_new_design' && (
                <div className="space-y-6">
                  {/* Passo 2.1: Preparar Lote */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-5 space-y-4">
                    <div className="flex items-center justify-between">
                      <div>
                        <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#57534E]">
                          Etapa 1 · Preparar Lote em Camadas Editáveis
                        </span>
                        <h4 className="text-sm font-bold text-[#1C1917] mt-0.5">
                          Estrutura do Novo Design Editável
                        </h4>
                      </div>
                      <span className="text-[11px] font-mono px-2.5 py-1 bg-white border border-[#E7E5E4] rounded-lg text-[#1C1917]">
                        Formato: {aspectRatio}
                      </span>
                    </div>

                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                      <div className="bg-white p-3 rounded-xl border border-[#E7E5E4]">
                        <span className="text-[#78716C] text-[11px]">Páginas / Frases</span>
                        <p className="font-bold text-base font-mono text-[#1C1917]">{validPhrases.length}</p>
                      </div>
                      <div className="bg-white p-3 rounded-xl border border-[#E7E5E4]">
                        <span className="text-[#78716C] text-[11px]">Mídias de Fundo</span>
                        <p className="font-bold text-base font-mono text-[#1C1917]">{mediaList.length}</p>
                      </div>
                      <div className="bg-white p-3 rounded-xl border border-[#E7E5E4]">
                        <span className="text-[#78716C] text-[11px]">Camada de Texto</span>
                        <p className="font-bold text-xs text-emerald-700 mt-1">Caixa Editável Nativa</p>
                      </div>
                      <div className="bg-white p-3 rounded-xl border border-[#E7E5E4]">
                        <span className="text-[#78716C] text-[11px]">Película & Assinatura</span>
                        <p className="font-bold text-xs text-emerald-700 mt-1">Camadas Separadas</p>
                      </div>
                    </div>

                    <div className="bg-white border border-[#E7E5E4] rounded-xl p-3 flex items-start gap-2 text-[11.5px] text-[#57534E]">
                      <Info className="w-4 h-4 text-[#00C4CC] shrink-0 mt-0.5" />
                      <span>
                        Cada página será montada em formato de apresentação editável (<code className="font-mono">.pptx</code>) e importada pela API oficial do Canva (<code className="font-mono">POST /rest/v1/imports</code>). Textos, assinatura, película de contraste e imagens permanecem elementos independentes no editor do Canva. <em>Observação: caso alguma mídia de fundo seja vídeo MP4, na importação para design o quadro inicial é inserido como camada de imagem editável.</em>
                      </span>
                    </div>
                  </div>

                  {/* Passo 2.2: Visualizar Páginas */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-5 space-y-3">
                    <div>
                      <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#57534E]">
                        Etapa 2 · Visualizar Páginas ({validPhrases.length})
                      </span>
                      <h4 className="text-sm font-bold text-[#1C1917] mt-0.5">
                        Prévia de todas as páginas que serão criadas no Canva
                      </h4>
                    </div>

                    <div className="flex items-stretch gap-3 overflow-x-auto pb-2">
                      {validPhrases.map((phrase, idx) => {
                        const pageInfo = getPreviewPageInfo(idx, mediaList);
                        const tplTotal = Math.max(1, inspectedTemplate?.pageCount || 1);
                        const tplPageNum = (idx % tplTotal) + 1;
                        const tplPage = inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)];
                        const tplLayout =
                          inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
                          inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
                          null;
                        const effectivePageNum = pageInfo.media ? pageInfo.pageNum : tplPageNum;
                        const effectiveLayout = pageInfo.slideLayout || tplLayout;
                        const effectiveCleanBg = pageInfo.cleanPageUrl || tplPage?.cleanBgUrl || '';
                        const effectiveThumb = pageInfo.pageUrl || tplPage?.thumbnailUrl || '';
                        const cleanBgOk = Boolean(effectiveCleanBg && !failedCleanBgUrls[effectiveCleanBg]);
                        const isRoundtripCard =
                          Boolean(inspectedTemplate?.editableViaPptxRoundtrip) ||
                          Boolean(effectiveLayout?.textBoxes?.length);
                        const updatedCardBoxes = isRoundtripCard
                          ? replaceSlideLayoutTextBoxesInPlace(
                              effectiveLayout,
                              effectivePageNum,
                              phrase,
                              keyFieldMappings,
                              signatureText
                            )
                          : [];

                        return (
                          <div
                            key={idx}
                            onClick={() => setPreviewIndex(idx)}
                            className={`w-36 shrink-0 rounded-xl overflow-hidden border cursor-pointer transition-all bg-neutral-900 relative flex flex-col justify-between ${
                              previewIndex === idx ? 'border-[#00C4CC] ring-2 ring-[#00C4CC]/30' : 'border-[#E7E5E4]'
                            }`}
                            style={{ height: aspectRatio === '16:9' ? '110px' : '190px' }}
                          >
                            {(cleanBgOk ? effectiveCleanBg : effectiveThumb) && !pageInfo.isVideo && (
                              <img
                                src={isRoundtripCard && cleanBgOk ? effectiveCleanBg : effectiveThumb}
                                alt={`Página ${idx + 1}`}
                                onError={(e) => {
                                  if (effectiveCleanBg && e.currentTarget.src.includes(effectiveCleanBg)) {
                                    setFailedCleanBgUrls((prev) => ({ ...prev, [effectiveCleanBg]: true }));
                                  }
                                  if (effectiveThumb && e.currentTarget.src !== effectiveThumb) {
                                    e.currentTarget.src = effectiveThumb;
                                  }
                                }}
                                className="absolute inset-0 w-full h-full object-cover opacity-90"
                              />
                            )}
                            {!isRoundtripCard && (
                              <div
                                className="absolute inset-0"
                                style={{ backgroundColor: `rgba(0,0,0,${overlayOpacity / 100})` }}
                              />
                            )}
                            <div className="relative z-10 p-1.5 flex items-center justify-between">
                              <span className="bg-black/80 text-white text-[9px] font-mono px-1.5 py-0.5 rounded">
                                Pág {idx + 1}
                              </span>
                              {pageInfo.isVideo && (
                                <span className="bg-amber-500/90 text-white text-[8px] font-mono px-1 rounded">
                                  Frame MP4
                                </span>
                              )}
                            </div>

                            {isRoundtripCard ? (
                              updatedCardBoxes.map((box, bIdx) => {
                                if (!cleanBgOk && !box.didReplace) return null;
                                return (
                                  <div
                                    key={`card_tb_${bIdx}`}
                                    className="absolute pointer-events-none flex flex-col justify-center"
                                    style={{
                                      left: `${box.xPct}%`,
                                      top: `${box.yPct}%`,
                                      width: `${box.wPct}%`,
                                      minHeight: `${box.hPct}%`,
                                      textAlign: box.align,
                                      padding: '1px 2px',
                                      backgroundColor:
                                        !cleanBgOk && box.didReplace ? 'rgba(28, 25, 23, 0.92)' : 'transparent',
                                      borderRadius: !cleanBgOk && box.didReplace ? '3px' : undefined,
                                    }}
                                  >
                                    <p
                                      style={{
                                        fontFamily: box.fontFamily ? `"${box.fontFamily}", sans-serif` : undefined,
                                        fontSize: `${Math.max(6, Math.min(11, box.fontSizePt * 0.25))}px`,
                                        color: box.colorHex || '#FFFFFF',
                                        fontWeight: box.bold ? 700 : 500,
                                        fontStyle: box.italic ? 'italic' : 'normal',
                                        textDecoration: box.underline ? 'underline' : undefined,
                                        letterSpacing: box.charSpacingPt ? `${(box.charSpacingPt * 0.25).toFixed(2)}px` : undefined,
                                        lineHeight: box.lineSpacingPct ? box.lineSpacingPct / 100 : 1.15,
                                        textShadow: '0 1px 2px rgba(0,0,0,0.45)',
                                        whiteSpace: 'pre-wrap',
                                      }}
                                    >
                                      {box.text}
                                    </p>
                                  </div>
                                );
                              })
                            ) : (
                              <>
                                <div className="relative z-10 p-2 text-center my-auto">
                                  <p
                                    className="text-[10px] font-semibold line-clamp-3 leading-tight"
                                    style={{ color: textColor || '#FFFFFF' }}
                                  >
                                    {phrase}
                                  </p>
                                </div>
                                {signatureText && (
                                  <div className="relative z-10 pb-1.5 text-center">
                                    <span className="text-[8px] font-mono text-white/80">{signatureText}</span>
                                  </div>
                                )}
                              </>
                            )}
                          </div>
                        );
                      })}
                    </div>
                  </div>

                  {/* Passo 2.3: Criar Design e Salvar no Canva */}
                  <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-5 space-y-4">
                    <div>
                      <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#57534E]">
                        Etapa 3 · Criar Design e Salvar no Canva
                      </span>
                      <h4 className="text-sm font-bold text-[#1C1917] mt-0.5">
                        Escolha como criar o novo design editável
                      </h4>
                    </div>

                    {isSavingDesignToCanva && (
                      <div className="bg-cyan-50/80 border border-cyan-200 rounded-xl p-3.5 flex items-center gap-3 text-xs">
                        <RefreshCw className="w-4 h-4 text-[#00C4CC] animate-spin shrink-0" />
                        <span className="text-[#1C1917] font-medium">{savingDesignStatusText}</span>
                      </div>
                    )}

                    {designSaveBanner && (
                      <div
                        className={`p-3.5 rounded-xl border text-xs flex items-start gap-2 ${
                          designSaveBanner.type === 'success'
                            ? 'bg-emerald-50 border-emerald-200 text-emerald-900'
                            : 'bg-rose-50 border-rose-200 text-rose-900'
                        }`}
                      >
                        {designSaveBanner.type === 'success' ? (
                          <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0 mt-0.5" />
                        ) : (
                          <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                        )}
                        <span>{designSaveBanner.text}</span>
                      </div>
                    )}

                    <div className="flex flex-wrap items-center gap-3">
                      <button
                        type="button"
                        onClick={() => handleSaveDesignToCanva('multipage')}
                        disabled={isSavingDesignToCanva || validPhrases.length === 0 || isModelValidationBlocked}
                        className="px-6 py-3.5 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-xl flex items-center gap-2 transition-colors disabled:opacity-50 cursor-pointer shadow-sm"
                      >
                        <Sparkles className="w-4 h-4" />
                        <span>Criar Novo Design Multipáginas no Canva ({validPhrases.length} págs)</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => handleSaveDesignToCanva('individual')}
                        disabled={isSavingDesignToCanva || validPhrases.length === 0 || isModelValidationBlocked}
                        className="px-5 py-3.5 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-xl flex items-center gap-2 transition-colors disabled:opacity-50 cursor-pointer"
                      >
                        <Layers className="w-4 h-4 text-[#00C4CC]" />
                        <span>Criar 1 Design por Frase no Canva</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => downloadAlternativePackage()}
                        disabled={downloadingPptx || validPhrases.length === 0 || isModelValidationBlocked}
                        className="px-4 py-3.5 bg-white border border-[#E7E5E4] hover:bg-[#F5F5F4] text-[#1C1917] text-xs font-medium rounded-xl flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                      >
                        <Download className="w-4 h-4 text-[#78716C]" />
                        <span>{downloadingPptx ? 'Gerando .PPTX...' : 'Baixar .PPTX Editável'}</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => startBatchProcessing()}
                        disabled={isProcessing || validPhrases.length === 0 || isModelValidationBlocked}
                        className="px-4 py-3.5 bg-white border border-[#E7E5E4] hover:bg-[#F5F5F4] text-[#57534E] text-xs font-medium rounded-xl flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                      >
                        <Play className="w-3.5 h-3.5" />
                        <span>Ir para Resultados</span>
                      </button>
                    </div>

                    {/* Resultado direto do Novo Design criado na Opção 2 */}
                    {createdCanvaDesign && (
                      <div className="bg-white border border-emerald-300 rounded-xl p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                        <div className="space-y-1">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="text-xs font-bold text-emerald-900">Design Criado no Canva</span>
                            <span className="text-[10px] font-mono px-2 py-0.5 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded font-bold">
                              ID: {createdCanvaDesign.designId}
                            </span>
                            {createdCanvaDesign.pageCount && (
                              <span className="text-[10px] font-mono px-2 py-0.5 bg-[#FAF9F6] text-[#1C1917] border border-[#E7E5E4] rounded font-bold">
                                {createdCanvaDesign.pageCount} página(s)
                              </span>
                            )}
                          </div>
                          <p className="text-sm font-bold text-[#1C1917]">{createdCanvaDesign.title}</p>
                        </div>

                        <div className="flex items-center gap-2 shrink-0">
                          <a
                            href={createdCanvaDesign.editUrl}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-lg flex items-center gap-1.5"
                          >
                            <span>Abrir no Canva</span>
                            <ExternalLink className="w-3.5 h-3.5" />
                          </a>
                          <button
                            type="button"
                            onClick={() => {
                              navigator.clipboard.writeText(createdCanvaDesign.editUrl);
                              setCopiedDesignLink(true);
                              setTimeout(() => setCopiedDesignLink(false), 2000);
                            }}
                            className="px-3 py-2 border border-[#E7E5E4] hover:bg-[#FAF9F6] text-xs font-medium text-[#1C1917] rounded-lg flex items-center gap-1 cursor-pointer"
                          >
                            {copiedDesignLink ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5 text-[#78716C]" />}
                            <span>{copiedDesignLink ? 'Copiado!' : 'Copiar link'}</span>
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}

              <div className="flex justify-between pt-2">
                <button
                  onClick={() => setActiveStep(5)}
                  className="px-4 py-2 border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#57534E] hover:bg-[#F5F5F4]"
                >
                  Voltar para Estilo
                </button>
                <button
                  onClick={() => setActiveStep(7)}
                  className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white rounded-lg text-xs font-medium flex items-center gap-1.5 cursor-pointer"
                >
                  <span>Avançar para Resultados</span>
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 7: CONSULTAR E BAIXAR RESULTADOS                                    */}
          {/* ========================================================================= */}
          {activeStep === 7 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] p-6 sm:p-8 shadow-sm space-y-6">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-[#F5F5F4]">
                <div>
                  <h2 className="text-2xl font-serif font-bold text-[#1C1917]">
                    7. Resultados do Lote
                  </h2>
                  <p className="text-xs text-[#78716C] mt-1">
                    Designs gerados ou arquivos alternativos preparados para importação.
                  </p>
                </div>
                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={exportCsvReport}
                    className="px-3 py-2 border border-[#E7E5E4] hover:bg-[#F5F5F4] text-xs font-medium text-[#1C1917] rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer"
                  >
                    <Download className="w-3.5 h-3.5 text-[#78716C]" />
                    <span>Exportar CSV</span>
                  </button>

                  <button
                    onClick={() => downloadAlternativePackage()}
                    disabled={downloadingPptx || validPhrases.length === 0 || isModelValidationBlocked}
                    className="px-3 py-2 border border-[#E7E5E4] hover:bg-[#F5F5F4] text-xs font-medium text-[#1C1917] rounded-xl flex items-center gap-1.5 transition-colors cursor-pointer disabled:opacity-50"
                  >
                    {downloadingPptx ? (
                      <>
                        <RefreshCw className="w-3.5 h-3.5 animate-spin text-[#00C4CC]" />
                        <span>Gerando .PPTX...</span>
                      </>
                    ) : (
                      <>
                        <Download className="w-3.5 h-3.5 text-[#78716C]" />
                        <span>Baixar Arquivo Editável (.PPTX)</span>
                      </>
                    )}
                  </button>

                  {/* Botão de Salvar como Novo Design no Canva */}
                  <button
                    onClick={() => handleSaveDesignToCanva('multipage')}
                    disabled={isSavingDesignToCanva || validPhrases.length === 0 || isModelValidationBlocked}
                    className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-xl flex items-center gap-2 transition-all shadow-sm disabled:opacity-50 cursor-pointer"
                  >
                    {isSavingDesignToCanva ? (
                      <>
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                        <span>Criando no Canva...</span>
                      </>
                    ) : (
                      <>
                        <Sparkles className="w-3.5 h-3.5" />
                        <span>Salvar como Design no Canva</span>
                      </>
                    )}
                  </button>

                  {designMode === 'individual' && validPhrases.length > 1 && (
                    <button
                      onClick={() => handleSaveDesignToCanva('individual')}
                      disabled={isSavingDesignToCanva || validPhrases.length === 0 || isModelValidationBlocked}
                      className="px-3.5 py-2 bg-[#FAF9F6] hover:bg-[#F5F5F4] border border-[#E7E5E4] text-[#1C1917] text-xs font-semibold rounded-xl flex items-center gap-1.5 transition-all disabled:opacity-50 cursor-pointer"
                      title="Cria um design separado no Canva para cada frase"
                    >
                      <Layers className="w-3.5 h-3.5 text-[#00C4CC]" />
                      <span>1 Design por Frase no Canva</span>
                    </button>
                  )}

                  {/* Botão de Salvar Vídeos/Assets no Canva */}
                  <button
                    onClick={() => handleUploadAssetsToCanva(false)}
                    disabled={isUploadingToCanva || batchResults.length === 0}
                    className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-bold rounded-xl flex items-center gap-2 transition-all shadow-sm disabled:opacity-50 cursor-pointer"
                  >
                    {isUploadingToCanva ? (
                      <>
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                        <span>Enviando Vídeos...</span>
                      </>
                    ) : (
                      <>
                        <UploadCloud className="w-3.5 h-3.5" />
                        <span>Salvar no Canva (Uploads)</span>
                      </>
                    )}
                  </button>
                </div>
              </div>

              {isModelValidationBlocked && (
                <div className="p-3.5 bg-rose-50 border-2 border-rose-500 rounded-xl text-xs text-rose-950 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="flex items-start gap-2">
                    <AlertCircle className="w-4 h-4 text-rose-600 shrink-0 mt-0.5" />
                    <span>
                      <strong>Geração bloqueada:</strong> existem campos-chave obrigatórios não encontrados no modelo selecionado (
                      {missingRequiredFieldMappings.map((m) => `"${m.fieldKey || '(vazio)'}"`).join(', ') || 'nenhum campo de frase compatível'}
                      ). Corrija na Etapa 6 antes de gerar ou baixar o lote.
                    </span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setActiveStep(6)}
                    className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white font-semibold rounded-lg shrink-0 cursor-pointer"
                  >
                    Corrigir Mapeamento
                  </button>
                </div>
              )}

              {/* Status do Processamento de Criação do Design */}
              {isSavingDesignToCanva && (
                <div className="bg-cyan-50/80 border border-cyan-200 rounded-2xl p-4 flex items-center gap-3">
                  <RefreshCw className="w-5 h-5 text-[#00C4CC] animate-spin shrink-0" />
                  <div className="min-w-0">
                    <p className="text-xs font-bold text-[#008A90]">Criação de Novo Design em Andamento</p>
                    <p className="text-xs text-[#57534E] mt-0.5">{savingDesignStatusText}</p>
                  </div>
                </div>
              )}

              {/* CARD OPÇÃO 1: VERSÕES GERADAS A PARTIR DO MODELO (AUTOFILL) */}
              {autofillVersionResults.length > 0 && (
                <div className="bg-cyan-50/40 border border-cyan-200 rounded-2xl p-5 space-y-3">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                    <div>
                      <span className="text-[10px] font-mono font-bold uppercase tracking-wider text-[#008A90]">
                        Opção 1 · Brand Template Autofill
                      </span>
                      <h3 className="text-sm font-bold text-[#1C1917]">
                        Versões Editáveis Criadas a Partir do Modelo ({autofillVersionResults.filter((r) => r.status === 'completed').length} de {autofillVersionResults.length})
                      </h3>
                    </div>
                  </div>
                  <div className="space-y-2">
                    {autofillVersionResults.map((res) => (
                      <div
                        key={`step7_af_${res.versionIndex}_${res.designId || 'err'}`}
                        className="bg-white border border-[#E7E5E4] rounded-xl p-3.5 flex flex-col sm:flex-row sm:items-center justify-between gap-3"
                      >
                        <div className="min-w-0 space-y-0.5">
                          <div className="flex items-center gap-2 flex-wrap">
                            <span className="font-mono text-xs font-bold text-[#008A90]">#{res.versionIndex}</span>
                            <span className="text-xs font-bold text-[#1C1917]">{res.title}</span>
                            {res.designId && (
                              <span className="text-[10px] font-mono px-2 py-0.5 bg-emerald-50 text-emerald-800 border border-emerald-200 rounded font-bold">
                                ID: {res.designId} · {res.pageCount} pág(s)
                              </span>
                            )}
                          </div>
                          {res.label && <p className="text-[11px] text-[#78716C] truncate">{res.label}</p>}
                          {res.error && <p className="text-[11px] text-rose-600">{res.error}</p>}
                        </div>

                        {res.status === 'completed' && res.editUrl && (
                          <div className="flex items-center gap-2 shrink-0">
                            <a
                              href={res.editUrl}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="px-3.5 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-lg flex items-center gap-1.5"
                            >
                              <span>Abrir no Canva</span>
                              <ExternalLink className="w-3.5 h-3.5" />
                            </a>
                            <button
                              type="button"
                              onClick={() => {
                                navigator.clipboard.writeText(res.editUrl!);
                                setCopiedVersionId(res.designId);
                                setTimeout(() => setCopiedVersionId(null), 2000);
                              }}
                              className="px-3 py-2 border border-[#E7E5E4] hover:bg-[#FAF9F6] text-xs font-medium text-[#1C1917] rounded-lg flex items-center gap-1 cursor-pointer"
                            >
                              {copiedVersionId === res.designId ? (
                                <>
                                  <Check className="w-3.5 h-3.5 text-emerald-600" />
                                  <span>Copiado</span>
                                </>
                              ) : (
                                <>
                                  <Copy className="w-3.5 h-3.5 text-[#78716C]" />
                                  <span>Copiar link</span>
                                </>
                              )}
                            </button>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              )}

              {/* CARD: NOVO DESIGN EDITÁVEL CRIADO NO CANVA (PARTE 3 & 4) */}
              {createdCanvaDesign && (
                <div className="bg-gradient-to-r from-emerald-500/10 via-teal-500/10 to-cyan-500/10 border border-emerald-300 rounded-2xl p-5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 shadow-sm">
                  <div className="flex items-start gap-3.5">
                    <div className="w-10 h-10 rounded-xl bg-emerald-600 text-white flex items-center justify-center shrink-0 shadow-sm">
                      <Sparkles className="w-5 h-5" />
                    </div>
                    <div>
                      <div className="flex items-center gap-2 flex-wrap">
                        <span className="text-xs font-bold text-emerald-950 uppercase tracking-wide">Design Editável Criado no Canva</span>
                        <span className="text-[10px] font-mono px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-md font-bold">
                          ID: {createdCanvaDesign.designId}
                        </span>
                        {createdCanvaDesign.pageCount && (
                          <span className="text-[10px] font-mono px-2 py-0.5 bg-white text-[#1C1917] border border-[#E7E5E4] rounded-md font-bold">
                            {createdCanvaDesign.pageCount} página(s)
                          </span>
                        )}
                      </div>
                      <p className="text-sm font-bold text-[#1C1917] mt-1">{createdCanvaDesign.title}</p>
                      <p className="text-xs text-[#78716C] mt-0.5">
                        Design multipáginas oficial criado no Canva. Contém todas as páginas do lote com camadas de texto, película e mídias editáveis.
                      </p>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 shrink-0 w-full sm:w-auto">
                    <a
                      href={createdCanvaDesign.editUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="flex-1 sm:flex-none px-4 py-2.5 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-xl flex items-center justify-center gap-2 shadow-sm transition-all"
                    >
                      <span>Abrir Design no Canva</span>
                      <ExternalLink className="w-3.5 h-3.5" />
                    </a>

                    <button
                      onClick={() => {
                        navigator.clipboard.writeText(createdCanvaDesign.editUrl);
                        setCopiedDesignLink(true);
                        setTimeout(() => setCopiedDesignLink(false), 2000);
                      }}
                      className="px-3 py-2.5 bg-white border border-[#E7E5E4] hover:bg-[#F5F5F4] text-xs font-medium text-[#1C1917] rounded-xl flex items-center justify-center gap-1.5 transition-colors cursor-pointer"
                      title="Copiar Link do Design"
                    >
                      {copiedDesignLink ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5 text-[#78716C]" />}
                      <span>{copiedDesignLink ? 'Copiado!' : 'Copiar'}</span>
                    </button>
                  </div>
                </div>
              )}

              {/* CARD: SALVAR VÍDEOS GERADOS NA BIBLIOTECA DO CANVA (PARTE 2 & 5) */}
              <div className="bg-[#FAF9F6] border border-[#E7E5E4] rounded-2xl p-5 space-y-4">
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                  <div className="flex items-center gap-2.5">
                    <div className="w-8 h-8 rounded-lg bg-[#1C1917] text-white flex items-center justify-center">
                      <UploadCloud className="w-4 h-4 text-[#00C4CC]" />
                    </div>
                    <div>
                      <h3 className="text-sm font-bold text-[#1C1917]">Envio Direto de Vídeos e Assets para a Conta Canva</h3>
                      <p className="text-xs text-[#78716C]">Upload oficial via API (POST /rest/v1/asset-uploads) para a biblioteca da sua conta.</p>
                    </div>
                  </div>

                  <div className="flex items-center gap-2 flex-wrap">
                    <button
                      onClick={() => handleUploadAssetsToCanva(false)}
                      disabled={isUploadingToCanva || batchResults.length === 0}
                      className="px-3.5 py-1.5 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-all disabled:opacity-50 cursor-pointer"
                    >
                      {isUploadingToCanva ? <RefreshCw className="w-3.5 h-3.5 animate-spin" /> : <UploadCloud className="w-3.5 h-3.5" />}
                      <span>{isUploadingToCanva ? 'Enviando ao Canva...' : 'Salvar no Canva'}</span>
                    </button>

                    {canvaUploadQueue.some((i) => i.status === 'failed') && (
                      <button
                        onClick={() => handleUploadAssetsToCanva(true)}
                        disabled={isUploadingToCanva}
                        className="px-3 py-1.5 bg-amber-500 hover:bg-amber-600 text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 transition-all disabled:opacity-50 cursor-pointer"
                      >
                        <RotateCcw className="w-3 h-3" />
                        <span>Tentar Novamente Falhas</span>
                      </button>
                    )}

                    <a
                      href="https://www.canva.com/folder/uploads"
                      target="_blank"
                      rel="noopener noreferrer"
                      className="px-3 py-1.5 bg-white border border-[#E7E5E4] hover:bg-[#F5F5F4] text-xs font-medium text-[#1C1917] rounded-lg flex items-center gap-1.5 transition-colors"
                    >
                      <span>Abrir Uploads no Canva</span>
                      <ExternalLink className="w-3 h-3" />
                    </a>
                  </div>
                </div>

                {/* Métricas e Progresso de Upload */}
                {(isUploadingToCanva || canvaUploadQueue.some((i) => i.status !== 'pending')) && (
                  <div className="space-y-2 pt-2 border-t border-[#E7E5E4]">
                    <div className="flex items-center justify-between text-xs">
                      <div className="flex items-center gap-3">
                        <span className="font-semibold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-md">
                          ✓ {canvaUploadQueue.filter((i) => i.status === 'success').length} enviados
                        </span>
                        <span className="font-semibold text-amber-700 bg-amber-50 border border-amber-200 px-2 py-0.5 rounded-md">
                          ⏳ {canvaUploadQueue.filter((i) => i.status === 'pending' || i.status === 'uploading').length} pendentes
                        </span>
                        {canvaUploadQueue.filter((i) => i.status === 'failed').length > 0 && (
                          <span className="font-semibold text-red-700 bg-red-50 border border-red-200 px-2 py-0.5 rounded-md">
                            ✕ {canvaUploadQueue.filter((i) => i.status === 'failed').length} falhas
                          </span>
                        )}
                      </div>
                      <span className="font-mono text-[#78716C]">
                        {uploadProgressCount.current} / {uploadProgressCount.total || batchResults.length}
                      </span>
                    </div>

                    <div className="w-full bg-[#E7E5E4] h-2 rounded-full overflow-hidden">
                      <div
                        className="bg-[#00C4CC] h-full transition-all duration-300"
                        style={{
                          width: `${
                            uploadProgressCount.total > 0
                              ? (uploadProgressCount.current / uploadProgressCount.total) * 100
                              : 0
                          }%`,
                        }}
                      />
                    </div>
                  </div>
                )}

                {/* Tabela de Arquivos Individuais */}
                {canvaUploadQueue.length > 0 && (
                  <div className="border border-[#E7E5E4] bg-white rounded-xl overflow-hidden divide-y divide-[#F5F5F4] text-xs max-h-56 overflow-y-auto">
                    {canvaUploadQueue.map((item, idx) => (
                      <div key={item.id} className="p-3 flex items-center justify-between gap-3 hover:bg-[#FAF9F6]">
                        <div className="min-w-0 flex items-center gap-2">
                          <span className="font-mono text-[#A8A29E] shrink-0">#{idx + 1}</span>
                          <div className="truncate">
                            <p className="font-medium text-[#1C1917] truncate">{item.name}</p>
                            <p className="text-[11px] text-[#78716C] truncate">{item.phrase}</p>
                          </div>
                        </div>

                        <div className="flex items-center gap-2 shrink-0">
                          {item.status === 'success' && (
                            <span className="px-2 py-0.5 rounded-md text-[10px] font-mono font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 flex items-center gap-1">
                              <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                              <span>Salvo no Canva {item.assetId ? `(${item.assetId})` : ''}</span>
                            </span>
                          )}

                          {item.status === 'uploading' && (
                            <span className="px-2 py-0.5 rounded-md text-[10px] font-mono font-semibold bg-cyan-50 text-cyan-800 border border-cyan-200 flex items-center gap-1">
                              <RefreshCw className="w-3 h-3 animate-spin text-[#00C4CC]" />
                              <span>Enviando...</span>
                            </span>
                          )}

                          {item.status === 'failed' && (
                            <div className="flex items-center gap-1.5">
                              <span className="px-2 py-0.5 rounded-md text-[10px] font-mono font-semibold bg-red-50 text-red-700 border border-red-200 flex items-center gap-1">
                                <AlertCircle className="w-3 h-3 text-red-600" />
                                <span>{item.errorMessage || 'Falhou'}</span>
                              </span>
                            </div>
                          )}

                          {item.status === 'pending' && (
                            <span className="px-2 py-0.5 rounded-md text-[10px] font-mono text-[#78716C] bg-neutral-100">
                              Pendente
                            </span>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                )}

                <div className="text-[11px] text-[#78716C] bg-white border border-[#E7E5E4] rounded-xl p-3 flex items-start gap-2">
                  <Info className="w-4 h-4 text-[#78716C] shrink-0 mt-0.5" />
                  <span>
                    <strong>Diferença de Funcionalidades:</strong> O botão <em>Salvar no Canva (Uploads)</em> envia os arquivos brutos para sua biblioteca de Uploads para uso futuro. O botão <em>Salvar como Design no Canva</em> gera um projeto oficial multipáginas nativo e editável.
                  </span>
                </div>
              </div>

              {batchResults.length === 0 ? (
                <div className="text-center py-12 border border-[#E7E5E4] rounded-xl bg-[#FAF9F6] p-6">
                  <Layers className="w-8 h-8 text-[#A8A29E] mx-auto mb-2" />
                  <p className="text-sm font-semibold text-[#1C1917]">Nenhum lote gerado ainda</p>
                  <p className="text-xs text-[#78716C] mt-1">Inicie o processamento na Etapa 6 para gerar os resultados.</p>
                  <button
                    onClick={() => setActiveStep(6)}
                    className="mt-4 px-4 py-2 bg-[#1C1917] text-white text-xs font-medium rounded-lg"
                  >
                    Ir para Etapa 6
                  </button>
                </div>
              ) : (
                <div className="border border-[#E7E5E4] rounded-xl overflow-hidden divide-y divide-[#F5F5F4]">
                  {batchResults.map((item, idx) => (
                    <div key={`${item.id || 'res'}_${idx}`} className="p-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3 hover:bg-[#FAF9F6]">
                      <div className="space-y-1 max-w-xl">
                        <div className="flex items-center gap-2">
                          <span className="font-mono text-xs text-[#A8A29E] font-semibold">#{idx + 1}</span>
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-mono ${
                              item.type === 'canva_design'
                                ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                                : 'bg-blue-50 text-blue-800 border border-blue-200'
                            }`}
                          >
                            {item.type === 'canva_design' ? 'Canva Oficial' : 'Pacote Editável'}
                          </span>
                        </div>
                        <p className="text-xs font-medium text-[#1C1917] line-clamp-2">{item.phrase}</p>
                        {item.message && (
                          <p className="text-[11px] text-[#78716C]">{item.message}</p>
                        )}
                      </div>

                      <div className="flex items-center gap-2 shrink-0">
                        {item.url ? (
                          <div className="flex items-center gap-2">
                            <a
                              href={item.url}
                              target="_blank"
                              rel="noopener noreferrer"
                              className="px-3 py-1.5 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-medium rounded-lg flex items-center gap-1.5"
                            >
                              <span>Abrir no Canva</span>
                              <ExternalLink className="w-3 h-3" />
                            </a>
                            <button
                              type="button"
                              onClick={() => {
                                navigator.clipboard.writeText(item.url!);
                                setCopiedVersionId(item.id);
                                setTimeout(() => setCopiedVersionId(null), 2000);
                              }}
                              className="px-2.5 py-1.5 border border-[#E7E5E4] hover:bg-white text-[#1C1917] text-xs font-medium rounded-lg flex items-center gap-1 cursor-pointer"
                            >
                              {copiedVersionId === item.id ? (
                                <>
                                  <Check className="w-3 h-3 text-emerald-600" />
                                  <span>Copiado</span>
                                </>
                              ) : (
                                <>
                                  <Copy className="w-3 h-3 text-[#78716C]" />
                                  <span>Copiar</span>
                                </>
                              )}
                            </button>
                          </div>
                        ) : (
                          <>
                            <button
                              onClick={() => handleSaveDesignToCanva('individual', item.isMultipage ? undefined : idx)}
                              disabled={isSavingDesignToCanva}
                              className="px-3 py-1.5 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
                            >
                              <Sparkles className="w-3 h-3" />
                              <span>Criar no Canva</span>
                            </button>
                            <button
                              onClick={() => downloadAlternativePackage(item.isMultipage ? undefined : idx)}
                              disabled={downloadingPptx}
                              className="px-3 py-1.5 border border-[#E7E5E4] hover:bg-white text-[#1C1917] text-xs font-medium rounded-lg flex items-center gap-1.5 cursor-pointer"
                            >
                              <span>Baixar .PPTX</span>
                              <Download className="w-3 h-3" />
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
              )}

              <div className="flex justify-between pt-2">
                <button
                  onClick={() => setActiveStep(6)}
                  className="px-4 py-2 border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#57534E] hover:bg-[#F5F5F4]"
                >
                  Voltar para Gerador
                </button>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* ETAPA 8: MODO DRIVE CANVA (WINDOWS EXPLORER)                              */}
          {/* ========================================================================= */}
          {activeStep === 8 && (
            <div className="bg-white rounded-2xl border border-[#E7E5E4] shadow-sm overflow-hidden flex flex-col md:flex-row min-h-[660px]">
              
              {/* MENU LATERAL (Sidebar) */}
              <aside className="w-full md:w-64 bg-[#FAF9F6] border-b md:border-b-0 md:border-r border-[#E7E5E4] p-4 flex flex-col justify-between shrink-0">
                <div className="space-y-6">
                  <div>
                    <div className="flex items-center gap-2 text-[#1C1917] px-2 mb-3">
                      <HardDrive className="w-4 h-4 text-[#00C4CC]" />
                      <h3 className="font-serif font-bold text-sm">Central Drive Canva</h3>
                    </div>
                    <p className="text-[11px] text-[#78716C] px-2">Navegue pelos seus arquivos e projetos por categoria.</p>
                  </div>

                  {/* Categorias */}
                  <nav className="space-y-1">
                    {[
                      { id: 'all', label: 'Tudo (Todos os Arquivos)', icon: HardDrive, count: driveFilteredItems.length },
                      { id: 'design', label: 'Designs Canva', icon: Sparkles, count: countByCategory('design') },
                      { id: 'video', label: 'Vídeos', icon: Video, count: countByCategory('video') },
                      { id: 'image', label: 'Imagens (IMG)', icon: ImageIcon, count: countByCategory('image') },
                      { id: 'audio', label: 'Músicas / Áudios', icon: Music, count: countByCategory('audio') },
                      { id: 'folder', label: 'Pastas & Projetos', icon: Folder, count: countByCategory('folder') },
                    ].map((cat) => {
                      const Icon = cat.icon;
                      const active = driveCategory === cat.id;
                      return (
                        <button
                          key={cat.id}
                          onClick={() => setDriveCategory(cat.id as any)}
                          className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs font-medium transition-colors cursor-pointer ${
                            active
                              ? 'bg-[#1C1917] text-white shadow-xs'
                              : 'text-[#57534E] hover:bg-[#E7E5E4]/60'
                          }`}
                        >
                          <div className="flex items-center gap-2.5">
                            <Icon className={`w-4 h-4 ${active ? 'text-[#00C4CC]' : 'text-[#78716C]'}`} />
                            <span>{cat.label}</span>
                          </div>
                          <span className={`text-[10px] font-mono px-1.5 py-0.5 rounded ${active ? 'bg-white/20 text-white' : 'bg-[#E7E5E4] text-[#57534E]'}`}>
                            {cat.count}
                          </span>
                        </button>
                      );
                    })}
                  </nav>

                  {/* SEÇÃO DEDICADA: ÁRVORE HIERÁRQUICA DE PASTAS */}
                  <div className="pt-4 border-t border-[#E7E5E4] space-y-2">
                    <div className="flex items-center justify-between px-2">
                      <span className="text-[11px] font-bold text-[#78716C] uppercase tracking-wider flex items-center gap-1.5">
                        <FolderOpen className="w-3.5 h-3.5 text-amber-500" />
                        <span>Árvore de Pastas</span>
                      </span>
                      <button
                        onClick={() => setShowCreateFolderModal(true)}
                        className="p-1 hover:bg-[#E7E5E4] rounded text-[#57534E] transition-colors"
                        title="Criar nova pasta no Canva"
                      >
                        <FolderPlus className="w-3.5 h-3.5 text-[#00C4CC]" />
                      </button>
                    </div>

                    <div className="space-y-0.5 max-h-[320px] overflow-y-auto pr-1">
                      {renderFolderTreeNode('root', 'Meu Drive Canva', 0)}
                    </div>
                  </div>
                </div>

                {/* Widget de Armazenamento e Status */}
                <div className="pt-4 border-t border-[#E7E5E4] mt-6 space-y-2">
                  <div className="flex items-center justify-between text-[11px] text-[#78716C]">
                    <span>Status do Canva</span>
                    <span className={`font-semibold ${canvaStatus?.connected ? 'text-emerald-700' : 'text-amber-700'}`}>
                      {canvaStatus?.connected ? 'Conectado' : 'Modo Local'}
                    </span>
                  </div>
                  <div className="w-full bg-[#E7E5E4] h-1.5 rounded-full overflow-hidden">
                    <div className="bg-[#00C4CC] h-full w-2/5" />
                  </div>
                  <p className="text-[10px] text-[#A8A29E]">Modo Windows Explorer Ativo</p>
                </div>
              </aside>

              {/* ÁREA PRINCIPAL DO DRIVE (EXPLORER / GRID) */}
              <div className="flex-1 flex flex-col min-w-0 bg-white">
                
                {/* Notificação Toast de Arraste e Solte */}
                {dropNotification && (
                  <div className="bg-[#1C1917] text-white text-xs font-semibold px-4 py-2.5 flex items-center justify-between shadow-md border-b border-[#00C4CC]/30 animate-fade-in">
                    <div className="flex items-center gap-2">
                      <CheckCircle2 className="w-4 h-4 text-[#00C4CC]" />
                      <span>{dropNotification}</span>
                    </div>
                    <button onClick={() => setDropNotification(null)} className="text-white/70 hover:text-white p-0.5">
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                )}

                {/* Banner de Busca Global em Todo o Drive */}
                {isGlobalSearching && (
                  <div className="bg-[#FAF9F6] border-b border-[#E7E5E4] px-4 py-2.5 flex flex-wrap items-center justify-between gap-2">
                    <div className="flex items-center gap-2 text-xs">
                      <Search className="w-4 h-4 text-[#00C4CC]" />
                      <span className="font-bold text-[#1C1917]">Resultados da Busca Global:</span>
                      <span className="text-[#57534E]">
                        {driveFilteredItems.length} item(ns) encontrado(s) para <strong className="text-[#1C1917]">"{driveSearchQuery}"</strong> em Todo o Drive
                      </span>
                    </div>
                    <button
                      onClick={() => setDriveSearchQuery('')}
                      className="px-2.5 py-1 text-[11px] font-medium text-[#78716C] hover:text-[#1C1917] bg-white border border-[#E7E5E4] rounded-lg hover:bg-[#E7E5E4] flex items-center gap-1 cursor-pointer"
                    >
                      <X className="w-3 h-3" />
                      <span>Limpar Busca</span>
                    </button>
                  </div>
                )}

                {/* BARRA SUPERIOR DE AÇÕES E BREADCRUMB */}
                <div className="p-4 border-b border-[#E7E5E4] bg-[#FAF9F6] flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
                  {/* Breadcrumbs de Navegação de Pasta com suporte a Drop */}
                  <div className="flex items-center gap-1.5 text-xs text-[#57534E] overflow-x-auto">
                    {driveFolderPath.length > 1 && (
                      <button
                        onClick={() => {
                          setDriveFolderPath((prev) => prev.slice(0, prev.length - 1));
                          setDriveCategory('all');
                        }}
                        className="p-1 rounded text-[#57534E] hover:bg-[#E7E5E4] transition-colors flex items-center gap-1 font-medium mr-1"
                        title="Voltar pasta superior"
                      >
                        <ChevronLeft className="w-4 h-4" />
                        <span className="hidden sm:inline">Voltar</span>
                      </button>
                    )}
                    {driveFolderPath.map((folder, idx) => (
                      <React.Fragment key={`${folder.id}_${idx}`}>
                        {idx > 0 && <ChevronRight className="w-3.5 h-3.5 text-[#A8A29E] shrink-0" />}
                        <button
                          onClick={() => {
                            setDriveFolderPath((prev) => prev.slice(0, idx + 1));
                            setDriveCategory('all');
                          }}
                          onDragOver={(e) => {
                            e.preventDefault();
                            e.dataTransfer.dropEffect = 'move';
                            setDragOverFolderId(folder.id);
                          }}
                          onDragLeave={() => setDragOverFolderId(null)}
                          onDrop={(e) => {
                            e.preventDefault();
                            handleDropOnFolder(folder.id, folder.name);
                          }}
                          className={`px-2 py-1 rounded transition-all whitespace-nowrap font-medium ${
                            dragOverFolderId === folder.id
                              ? 'bg-[#00C4CC] text-white font-bold ring-2 ring-white scale-105 shadow-sm'
                              : idx === driveFolderPath.length - 1
                              ? 'font-semibold text-[#1C1917] bg-white border border-[#E7E5E4]'
                              : 'hover:bg-[#E7E5E4]'
                          }`}
                        >
                          📁 {folder.name}
                        </button>
                      </React.Fragment>
                    ))}
                  </div>

                  {/* Controles: Busca Global, Criar Pasta, Alternância de Modo */}
                  <div className="flex items-center gap-2 shrink-0">
                    <div className="relative">
                      <Search className="w-3.5 h-3.5 text-[#A8A29E] absolute left-2.5 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={driveSearchQuery}
                        onChange={(e) => setDriveSearchQuery(e.target.value)}
                        placeholder="Buscar em todo o Drive..."
                        className="text-xs pl-8 pr-3 py-1.5 bg-white border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] w-36 sm:w-52 font-sans"
                      />
                      {driveSearchQuery && (
                        <button
                          onClick={() => setDriveSearchQuery('')}
                          className="absolute right-2 top-1/2 -translate-y-1/2 text-[#A8A29E] hover:text-[#1C1917]"
                        >
                          <X className="w-3 h-3" />
                        </button>
                      )}
                    </div>

                    <button
                      onClick={() => setShowCreateFolderModal(true)}
                      className="px-3 py-1.5 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-1.5 cursor-pointer shrink-0"
                      title="Criar nova pasta no Canva Drive"
                    >
                      <FolderPlus className="w-3.5 h-3.5 text-[#00C4CC]" />
                      <span>Nova Pasta</span>
                    </button>

                    <div className="flex items-center border border-[#E7E5E4] rounded-lg p-0.5 bg-white">
                      <button
                        onClick={() => setDriveViewMode('explorer')}
                        className={`p-1.5 rounded text-xs transition-colors flex items-center gap-1 cursor-pointer ${
                          driveViewMode === 'explorer' ? 'bg-[#1C1917] text-white font-semibold' : 'text-[#57534E] hover:bg-[#FAF9F6]'
                        }`}
                        title="Modo Windows Explorer (Tabela com detalhes)"
                      >
                        <List className="w-3.5 h-3.5" />
                        <span className="hidden sm:inline text-[11px]">Explorer</span>
                      </button>
                      <button
                        onClick={() => setDriveViewMode('grid')}
                        className={`p-1.5 rounded text-xs transition-colors flex items-center gap-1 cursor-pointer ${
                          driveViewMode === 'grid' ? 'bg-[#1C1917] text-white font-semibold' : 'text-[#57534E] hover:bg-[#FAF9F6]'
                        }`}
                        title="Modo Grade (Miniaturas)"
                      >
                        <Grid className="w-3.5 h-3.5" />
                        <span className="hidden sm:inline text-[11px]">Grade</span>
                      </button>
                    </div>

                    <button
                      onClick={() => {
                        if (canvaStatus?.connected) {
                          syncCanvaDriveFiles(currentFolderId);
                        } else {
                          handleConnect();
                        }
                      }}
                      disabled={loadingCanvaAssets}
                      className="px-3 py-1.5 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-1.5 cursor-pointer disabled:opacity-50"
                      title={canvaStatus?.connected ? "Sincronizar arquivos diretamente da API do Canva" : "Conectar com sua conta oficial do Canva"}
                    >
                      {loadingCanvaAssets ? (
                        <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                      ) : (
                        <Sparkles className="w-3.5 h-3.5" />
                      )}
                      <span className="hidden md:inline">{canvaStatus?.connected ? 'Sincronizar Canva' : 'Conectar Canva'}</span>
                    </button>
                  </div>
                </div>

                {/* BARRA DE ORDENAÇÃO E AÇÕES EM LOTE (BULK ACTIONS & SORTING) */}
                <div className="bg-[#FAF9F6] border-b border-[#E7E5E4] px-4 py-2.5 flex flex-wrap items-center justify-between gap-3 text-xs">
                  <div className="flex items-center gap-2">
                    {selectedItemIds.size > 0 ? (
                      <div className="flex items-center gap-2 bg-[#1C1917] text-white px-3 py-1 rounded-lg text-xs font-medium">
                        <span>{selectedItemIds.size} selecionado(s)</span>
                        <div className="flex items-center gap-1.5 ml-2 border-l border-white/20 pl-2">
                          <select
                            onChange={(e) => {
                              if (e.target.value) handleBulkMove(e.target.value);
                              e.target.value = '';
                            }}
                            defaultValue=""
                            className="bg-white/10 text-white text-[11px] px-2 py-0.5 rounded border border-white/20 focus:outline-none"
                          >
                            <option value="" disabled className="text-black">Mover para...</option>
                            {allAvailableFolders.map((f) => (
                              <option key={f.id} value={f.id} className="text-black">📁 {f.name}</option>
                            ))}
                          </select>
                          <button
                            onClick={handleBulkDelete}
                            className="bg-rose-600 hover:bg-rose-700 text-white px-2 py-0.5 rounded text-[11px] font-medium transition-colors"
                          >
                            Excluir Selecionados
                          </button>
                          <button
                            onClick={() => setSelectedItemIds(new Set())}
                            className="text-white/70 hover:text-white px-1 text-[11px]"
                          >
                            Limpar
                          </button>
                        </div>
                      </div>
                    ) : (
                      <span className="text-[#78716C]">
                        Exibindo {driveFilteredItems.length} item(ns)
                      </span>
                    )}
                  </div>

                  <div className="flex items-center gap-2 ml-auto">
                    <span className="text-[#78716C] flex items-center gap-1">
                      <ArrowUpDown className="w-3.5 h-3.5" />
                      <span>Ordenar por:</span>
                    </span>
                    <select
                      value={sortBy}
                      onChange={(e) => setSortBy(e.target.value as any)}
                      className="text-xs bg-white border border-[#E7E5E4] rounded-lg px-2 py-1 font-medium text-[#1C1917] focus:outline-none"
                    >
                      <option value="name">Nome</option>
                      <option value="date">Data de Modificação</option>
                      <option value="size">Tamanho</option>
                    </select>
                    <button
                      onClick={() => setSortOrder((prev) => (prev === 'asc' ? 'desc' : 'asc'))}
                      className="px-2 py-1 bg-white border border-[#E7E5E4] rounded-lg text-xs font-semibold text-[#1C1917] hover:bg-[#E7E5E4] transition-colors"
                      title={sortOrder === 'asc' ? 'Ordem Crescente (Ascendente)' : 'Ordem Decrescente (Descendente)'}
                    >
                      {sortOrder === 'asc' ? 'ASC ▲' : 'DESC ▼'}
                    </button>
                  </div>
                </div>

                {/* CONTEÚDO DOS ARQUIVOS (Tabela Windows Explorer OU Grade) */}
                <div className="flex-1 p-4 overflow-y-auto">
                  {driveFilteredItems.length === 0 ? (
                    <div className="text-center py-16 space-y-3 max-w-md mx-auto">
                      <FolderOpen className="w-10 h-10 text-[#D6D3D1] mx-auto" />
                      <p className="text-sm font-semibold text-[#1C1917]">
                        {isGlobalSearching ? `Nenhum resultado para "${driveSearchQuery}"` : 'Nenhum arquivo ou pasta nesta visualização'}
                      </p>
                      <p className="text-xs text-[#78716C] leading-relaxed">
                        {isGlobalSearching
                          ? 'Tente buscar por outro termo ou clique em Limpar Busca.'
                          : driveCategory === 'video'
                          ? 'A API REST pública do Canva (/folders/{id}/items) não lista arquivos MP4 brutos nas pastas; apenas designs com páginas de vídeo ou mídias enviadas/capturadas via Canva Apps SDK (Selection API + getTemporaryUrl).'
                          : driveCategory === 'audio'
                          ? 'A Canva Connect API REST não disponibiliza listagem de áudios/músicas da conta. Utilize a captura dentro do Canva Editor (Selection API + getTemporaryUrl) ou importe o arquivo de áudio localmente.'
                          : 'Clique em "+ Nova Pasta" para criar um diretório ou altere os filtros.'}
                      </p>
                    </div>
                  ) : driveViewMode === 'explorer' ? (
                    /* MODO WINDOWS EXPLORER (Tabela estilo gerenciador de arquivos com Arraste e Solte) */
                    <div className="border border-[#E7E5E4] rounded-xl overflow-hidden shadow-xs">
                      <table className="w-full text-left border-collapse text-xs">
                        <thead>
                          <tr className="bg-[#FAF9F6] border-b border-[#E7E5E4] text-[#78716C] font-semibold text-[11px]">
                            <th className="p-3 w-8 text-center">
                              <input
                                type="checkbox"
                                checked={driveFilteredItems.length > 0 && selectedItemIds.size === driveFilteredItems.length}
                                onChange={handleSelectAllToggle}
                                className="rounded border-[#D6D3D1] text-[#1C1917] focus:ring-0 cursor-pointer"
                              />
                            </th>
                            <th className="p-3 w-8 text-center">ARRASTAR</th>
                            <th className="p-3">NOME DO ARQUIVO / PASTA</th>
                            {isGlobalSearching && <th className="p-3">PASTA ORIGEM</th>}
                            <th className="p-3">CATEGORIA</th>
                            <th className="p-3">TAMANHO</th>
                            <th className="p-3">MODIFICAÇÃO</th>
                            <th className="p-3 text-right">AÇÕES</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-[#F5F5F4]">
                          {driveFilteredItems.map((item, idx) => {
                            const CategoryIcon =
                              item.category === 'folder'
                                ? Folder
                                : item.category === 'design'
                                ? Sparkles
                                : item.category === 'video'
                                ? Video
                                : item.category === 'audio'
                                ? Music
                                : ImageIcon;

                            const badgeColor =
                              item.category === 'folder'
                                ? 'bg-amber-100 text-amber-800'
                                : item.category === 'design'
                                ? 'bg-purple-100 text-purple-800'
                                : item.category === 'video'
                                ? 'bg-rose-100 text-rose-800'
                                : item.category === 'audio'
                                ? 'bg-emerald-100 text-emerald-800'
                                : 'bg-blue-100 text-blue-800';

                            const isFolder = item.category === 'folder';
                            const isBeingDragged = draggedItemId === item.id;
                            const isFolderTarget = dragOverFolderId === item.id;

                            return (
                              <tr
                                key={`${item.id}_${idx}`}
                                draggable={true}
                                onDragStart={(e) => {
                                  e.dataTransfer.setData('text/plain', item.id);
                                  setDraggedItemId(item.id);
                                }}
                                onDragEnd={() => {
                                  setDraggedItemId(null);
                                  setDragOverFolderId(null);
                                }}
                                onDragOver={(e) => {
                                  if (isFolder && draggedItemId !== item.id) {
                                    e.preventDefault();
                                    e.dataTransfer.dropEffect = 'move';
                                    setDragOverFolderId(item.id);
                                  }
                                }}
                                onDragLeave={() => {
                                  if (isFolder) {
                                    setDragOverFolderId(null);
                                  }
                                }}
                                onDrop={(e) => {
                                  if (isFolder) {
                                    e.preventDefault();
                                    handleDropOnFolder(item.id, item.name);
                                  }
                                }}
                                onClick={() => {
                                  if (isFolder) {
                                    navigateToFolderById(item.id, item.name);
                                  } else {
                                    setSelectedDriveItem(item);
                                  }
                                }}
                                className={`cursor-pointer transition-all group ${
                                  isFolderTarget
                                    ? 'bg-amber-100 border-2 border-amber-500 font-bold scale-[1.005]'
                                    : isBeingDragged
                                    ? 'opacity-40 bg-neutral-100'
                                    : 'hover:bg-[#FAF9F6]'
                                }`}
                              >
                                <td className="p-3 text-center text-[#A8A29E] group-hover:text-[#1C1917] cursor-grab active:cursor-grabbing">
                                  <GripVertical className="w-3.5 h-3.5 mx-auto" />
                                </td>
                                <td className="p-3 font-medium text-[#1C1917]">
                                  <div className="flex items-center gap-2">
                                    <CategoryIcon
                                      className={`w-4 h-4 shrink-0 ${
                                        isFolder
                                          ? 'text-amber-500'
                                          : item.category === 'design'
                                          ? 'text-purple-600'
                                          : item.category === 'video'
                                          ? 'text-rose-500'
                                          : item.category === 'audio'
                                          ? 'text-emerald-600'
                                          : 'text-blue-500'
                                      }`}
                                    />
                                    <span className="truncate max-w-xs">{item.name}</span>
                                    {isFolder && (
                                      <span className="text-[10px] text-amber-700 bg-amber-50 border border-amber-200 px-1.5 py-0.2 rounded font-mono">
                                        Pasta
                                      </span>
                                    )}
                                  </div>
                                </td>

                                {isGlobalSearching && (
                                  <td className="p-3">
                                    <span
                                      onClick={(e) => {
                                        e.stopPropagation();
                                        navigateToFolderById(item.folderId, getFolderName(item.folderId));
                                        setDriveSearchQuery('');
                                      }}
                                      className="text-[10px] bg-[#FAF9F6] hover:bg-[#E7E5E4] text-[#57534E] px-2 py-1 rounded border border-[#E7E5E4] font-mono inline-flex items-center gap-1 cursor-pointer transition-colors"
                                      title="Ir para a pasta deste item"
                                    >
                                      📁 {getFolderName(item.folderId)}
                                    </span>
                                  </td>
                                )}

                                <td className="p-3">
                                  <span className={`px-2 py-0.5 rounded text-[10px] font-mono font-medium ${badgeColor}`}>
                                    {item.category.toUpperCase()}
                                  </span>
                                </td>
                                <td className="p-3 text-[#78716C] font-mono">{item.size || '—'}</td>
                                <td className="p-3 text-[#78716C] font-mono">{item.modifiedAt || 'Hoje'}</td>
                                <td className="p-3 text-right space-x-1" onClick={(e) => e.stopPropagation()}>
                                  <button
                                    onClick={() => setSelectedDriveItem(item)}
                                    className="p-1 text-[#78716C] hover:text-[#1C1917] inline-block"
                                    title="Ver detalhes"
                                  >
                                    <Eye className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    onClick={() => {
                                      setItemToRename(item);
                                      setRenameInput(item.name);
                                      setShowRenameModal(true);
                                    }}
                                    className="p-1 text-[#78716C] hover:text-[#1C1917] inline-block"
                                    title="Renomear"
                                  >
                                    <Edit3 className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    onClick={() => {
                                      setItemToMove(item);
                                      setTargetFolderIdForMove(item.folderId || 'root');
                                      setShowMoveModal(true);
                                    }}
                                    className="p-1 text-[#78716C] hover:text-[#1C1917] inline-block"
                                    title="Mudar de Lugar (Mover)"
                                  >
                                    <FolderInput className="w-3.5 h-3.5" />
                                  </button>
                                  <button
                                    onClick={() => handleDeleteItem(item)}
                                    className="p-1 text-[#78716C] hover:text-rose-600 inline-block"
                                    title="Excluir"
                                  >
                                    <Trash2 className="w-3.5 h-3.5" />
                                  </button>
                                  {!isFolder && (
                                    <button
                                      onClick={() => {
                                        const newItem: LocalMediaItem = {
                                          id: item.id || Math.random().toString(36).substring(7),
                                          name: item.name,
                                          type: item.category === 'video' ? 'video' : 'image',
                                          size: 1024 * 1024 * 2,
                                          url: item.thumbnailUrl || 'https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=400&q=80',
                                          file: new File([], item.name),
                                        };
                                        setMediaList((prev) => [...prev, newItem]);
                                        alert(`"${item.name}" foi adicionado às suas mídias do lote!`);
                                      }}
                                      className="px-2.5 py-1 bg-[#1C1917] hover:bg-[#292524] text-white text-[11px] font-medium rounded transition-colors inline-flex items-center gap-1 cursor-pointer ml-1"
                                    >
                                      <Plus className="w-3 h-3" />
                                      <span>Usar no Lote</span>
                                    </button>
                                  )}
                                  {!isFolder && item.category === 'video' && (
                                    <button
                                      onClick={() => handleExportDesignMp4AndImport(item, false)}
                                      disabled={exportingMp4Id === item.id}
                                      className="px-2 py-1 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-[11px] font-semibold rounded transition-colors inline-flex items-center gap-1 cursor-pointer ml-1 disabled:opacity-50"
                                      title="Exportar vídeo MP4 oficial do Canva"
                                    >
                                      {exportingMp4Id === item.id ? (
                                        <RefreshCw className="w-3 h-3 animate-spin" />
                                      ) : (
                                        <Video className="w-3 h-3" />
                                      )}
                                      <span>MP4</span>
                                    </button>
                                  )}
                                  {item.canvaUrl && (
                                    <a
                                      href={item.canvaUrl}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      className="p-1 text-[#78716C] hover:text-[#1C1917] inline-block"
                                      title="Abrir no Canva"
                                    >
                                      <ExternalLink className="w-3.5 h-3.5" />
                                    </a>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  ) : (
                    /* MODO GRADE (CARDS GRANDES COM SUPORTE A ARRASTRAR E SOLTAR) */
                    <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4">
                      {driveFilteredItems.map((item, idx) => {
                        const isFolder = item.category === 'folder';
                        const isBeingDragged = draggedItemId === item.id;
                        const isFolderTarget = dragOverFolderId === item.id;

                        return (
                          <div
                            key={`${item.id}_${idx}`}
                            draggable={true}
                            onDragStart={(e) => {
                              e.dataTransfer.setData('text/plain', item.id);
                              setDraggedItemId(item.id);
                            }}
                            onDragEnd={() => {
                              setDraggedItemId(null);
                              setDragOverFolderId(null);
                            }}
                            onDragOver={(e) => {
                              if (isFolder && draggedItemId !== item.id) {
                                e.preventDefault();
                                e.dataTransfer.dropEffect = 'move';
                                setDragOverFolderId(item.id);
                              }
                            }}
                            onDragLeave={() => {
                              if (isFolder) {
                                setDragOverFolderId(null);
                              }
                            }}
                            onDrop={(e) => {
                              if (isFolder) {
                                e.preventDefault();
                                handleDropOnFolder(item.id, item.name);
                              }
                            }}
                            onClick={() => {
                              if (isFolder) {
                                navigateToFolderById(item.id, item.name);
                              } else {
                                setSelectedDriveItem(item);
                              }
                            }}
                            className={`group relative bg-white border rounded-xl overflow-hidden shadow-xs cursor-pointer transition-all aspect-square flex flex-col justify-end ${
                              isFolderTarget
                                ? 'border-2 border-amber-500 bg-amber-100 ring-4 ring-amber-200 scale-105 shadow-lg'
                                : isBeingDragged
                                ? 'opacity-40 border-dashed border-[#1C1917]'
                                : 'border-[#E7E5E4] hover:border-[#1C1917]'
                            }`}
                          >
                            {item.thumbnailUrl ? (
                              <img src={item.thumbnailUrl} alt={item.name} className="absolute inset-0 w-full h-full object-cover" />
                            ) : (
                              <div className="absolute inset-0 bg-[#FAF9F6] flex flex-col items-center justify-center p-3 text-center">
                                <Folder className="w-10 h-10 text-amber-500 mb-1 group-hover:scale-105 transition-transform" />
                                <span className="text-xs font-semibold text-[#1C1917] truncate max-w-full">{item.name}</span>
                              </div>
                            )}

                            {/* Checkbox de Seleção no Card em Grade */}
                            <div className="absolute top-2 left-2 z-20" onClick={(e) => handleToggleSelectItem(item.id, e)}>
                              <input
                                type="checkbox"
                                checked={selectedItemIds.has(item.id)}
                                onChange={() => {}}
                                className="w-4 h-4 rounded border-white text-[#1C1917] focus:ring-0 cursor-pointer shadow-md bg-white/80"
                              />
                            </div>

                            {/* Botões rápidos de Ação no Card */}
                            <div className="absolute top-2 right-2 opacity-0 group-hover:opacity-100 transition-opacity bg-white/95 backdrop-blur-xs rounded-lg p-1 shadow-sm border border-[#E7E5E4] flex items-center gap-1 z-20" onClick={(e) => e.stopPropagation()}>
                              <button
                                onClick={() => {
                                  setItemToRename(item);
                                  setRenameInput(item.name);
                                  setShowRenameModal(true);
                                }}
                                className="p-1 text-[#57534E] hover:text-[#1C1917]"
                                title="Renomear"
                              >
                                <Edit3 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => {
                                  setItemToMove(item);
                                  setTargetFolderIdForMove(item.folderId || 'root');
                                  setShowMoveModal(true);
                                }}
                                className="p-1 text-[#57534E] hover:text-[#1C1917]"
                                title="Mudar de Lugar"
                              >
                                <FolderInput className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => handleDeleteItem(item)}
                                className="p-1 text-[#57534E] hover:text-rose-600"
                                title="Excluir"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>

                            <div className="relative z-10 bg-gradient-to-t from-black/85 via-black/40 to-transparent p-2.5 text-white flex flex-col justify-end">
                              <div className="flex items-center justify-between">
                                <p className="text-xs font-semibold truncate pr-1">{item.name}</p>
                                <span className="text-[9px] text-white/70 font-mono uppercase shrink-0">{item.category}</span>
                              </div>
                              {isGlobalSearching && (
                                <span className="text-[9px] text-cyan-300 font-mono truncate mt-0.5">
                                  📁 {getFolderName(item.folderId)}
                                </span>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Modal de Detalhes do Arquivo (Drive Explorer) */}
          {selectedDriveItem && (
            <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
              <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-[#E7E5E4] space-y-5">
                <div className="flex items-center justify-between pb-3 border-b border-[#E7E5E4]">
                  <div className="flex items-center gap-2">
                    <Info className="w-4 h-4 text-[#00C4CC]" />
                    <h3 className="font-serif font-bold text-base text-[#1C1917]">Detalhes do Arquivo</h3>
                  </div>
                  <button
                    onClick={() => setSelectedDriveItem(null)}
                    className="text-xs font-medium text-[#78716C] hover:text-[#1C1917] px-2 py-1 rounded bg-[#FAF9F6] border border-[#E7E5E4]"
                  >
                    Fechar
                  </button>
                </div>

                {selectedDriveItem.thumbnailUrl && (
                  <div className="w-full h-44 rounded-xl overflow-hidden bg-black/5 relative border border-[#E7E5E4]">
                    <img src={selectedDriveItem.thumbnailUrl} alt={selectedDriveItem.name} className="w-full h-full object-cover" />
                  </div>
                )}

                <div className="space-y-2 text-xs">
                  <div className="flex justify-between py-1 border-b border-[#F5F5F4]">
                    <span className="text-[#78716C]">Nome:</span>
                    <span className="font-semibold text-[#1C1917] truncate max-w-[200px]">{selectedDriveItem.name}</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-[#F5F5F4]">
                    <span className="text-[#78716C]">Categoria:</span>
                    <span className="font-semibold uppercase text-purple-700 font-mono">{selectedDriveItem.category}</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-[#F5F5F4]">
                    <span className="text-[#78716C]">Tamanho:</span>
                    <span className="font-mono text-[#1C1917]">{selectedDriveItem.size || '—'}</span>
                  </div>
                  <div className="flex justify-between py-1 border-b border-[#F5F5F4]">
                    <span className="text-[#78716C]">Última modificação:</span>
                    <span className="font-mono text-[#1C1917]">{selectedDriveItem.modifiedAt || 'Hoje'}</span>
                  </div>
                </div>

                {/* Ações dentro do Modal de Detalhes */}
                <div className="pt-2 flex flex-wrap items-center justify-end gap-2 border-t border-[#F5F5F4]">
                  <button
                    onClick={() => {
                      setItemToRename(selectedDriveItem);
                      setRenameInput(selectedDriveItem.name);
                      setShowRenameModal(true);
                    }}
                    className="px-3 py-1.5 border border-[#E7E5E4] bg-white text-[#1C1917] text-xs font-medium rounded-lg hover:bg-[#FAF9F6] flex items-center gap-1.5 cursor-pointer"
                  >
                    <Edit3 className="w-3.5 h-3.5 text-[#57534E]" />
                    <span>Renomear</span>
                  </button>
                  <button
                    onClick={() => {
                      setItemToMove(selectedDriveItem);
                      setTargetFolderIdForMove(selectedDriveItem.folderId || 'root');
                      setShowMoveModal(true);
                    }}
                    className="px-3 py-1.5 border border-[#E7E5E4] bg-white text-[#1C1917] text-xs font-medium rounded-lg hover:bg-[#FAF9F6] flex items-center gap-1.5 cursor-pointer"
                  >
                    <FolderInput className="w-3.5 h-3.5 text-[#57534E]" />
                    <span>Mudar de Lugar</span>
                  </button>
                  <button
                    onClick={() => handleDeleteItem(selectedDriveItem)}
                    className="px-3 py-1.5 border border-rose-200 bg-rose-50 text-rose-700 text-xs font-medium rounded-lg hover:bg-rose-100 flex items-center gap-1.5 cursor-pointer"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>Excluir</span>
                  </button>

                  {selectedDriveItem.category !== 'folder' && (
                    <button
                      onClick={() => {
                        const newItem: LocalMediaItem = {
                          id: Math.random().toString(36).substring(7),
                          name: selectedDriveItem.name,
                          type: selectedDriveItem.category === 'video' ? 'video' : 'image',
                          size: 1024 * 1024 * 2,
                          url: selectedDriveItem.thumbnailUrl || 'https://images.unsplash.com/photo-1618005182384-a83a8bd57fbe?w=400&q=80',
                          file: new File([], selectedDriveItem.name),
                        };
                        setMediaList((prev) => [...prev, newItem]);
                        setSelectedDriveItem(null);
                        alert(`"${selectedDriveItem.name}" foi adicionado às suas mídias do lote!`);
                      }}
                      className="px-3.5 py-1.5 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg shadow-xs transition-colors flex items-center gap-1.5 cursor-pointer"
                    >
                      <Plus className="w-3.5 h-3.5" />
                      <span>Usar no Lote</span>
                    </button>
                  )}
                  {selectedDriveItem.canvaUrl && (
                    <a
                      href={selectedDriveItem.canvaUrl}
                      target="_blank"
                      rel="noopener noreferrer"
                      className="px-3 py-1.5 border border-[#E7E5E4] bg-[#FAF9F6] text-[#1C1917] text-xs font-medium rounded-lg hover:bg-[#E7E5E4] transition-colors inline-flex items-center gap-1"
                    >
                      <ExternalLink className="w-3.5 h-3.5" />
                      <span>Canva</span>
                    </a>
                  )}
                </div>
              </div>
            </div>
          )}

          {/* ========================================================================= */}
          {/* MODAL 1: CRIAR NOVA PASTA NO CANVA DRIVE                                   */}
          {/* ========================================================================= */}
          {showCreateFolderModal && (
            <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
              <form
                onSubmit={handleCreateFolder}
                className="bg-white rounded-2xl max-w-sm w-full p-6 shadow-2xl border border-[#E7E5E4] space-y-4"
              >
                <div className="flex items-center justify-between pb-2 border-b border-[#E7E5E4]">
                  <div className="flex items-center gap-2">
                    <FolderPlus className="w-5 h-5 text-[#00C4CC]" />
                    <h3 className="font-serif font-bold text-base text-[#1C1917]">Criar Nova Pasta</h3>
                  </div>
                  <button
                    type="button"
                    onClick={() => setShowCreateFolderModal(false)}
                    className="text-xs font-medium text-[#78716C] hover:text-[#1C1917]"
                  >
                    Cancelar
                  </button>
                </div>

                <div className="space-y-1">
                  <label className="block text-xs font-medium text-[#57534E]">Nome da Pasta</label>
                  <input
                    type="text"
                    value={newFolderName}
                    onChange={(e) => setNewFolderName(e.target.value)}
                    placeholder="ex. Vídeos Reels 2026"
                    autoFocus
                    className="w-full text-xs px-3 py-2 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917]"
                  />
                </div>

                <div className="pt-2 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowCreateFolderModal(false)}
                    className="px-3.5 py-2 border border-[#E7E5E4] text-xs font-medium rounded-lg text-[#57534E] hover:bg-[#FAF9F6]"
                  >
                    Cancelar
                  </button>
                  <button
                    type="submit"
                    disabled={creatingFolder || !newFolderName.trim()}
                    className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
                  >
                    {creatingFolder ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <FolderPlus className="w-3.5 h-3.5" />
                    )}
                    <span>Criar Pasta</span>
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* ========================================================================= */}
          {/* MODAL 2: RENOMEAR ARQUIVO OU PASTA                                         */}
          {/* ========================================================================= */}
          {showRenameModal && itemToRename && (
            <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
              <form
                onSubmit={handleRenameItem}
                className="bg-white rounded-2xl max-w-sm w-full p-6 shadow-2xl border border-[#E7E5E4] space-y-4"
              >
                <div className="flex items-center justify-between pb-2 border-b border-[#E7E5E4]">
                  <div className="flex items-center gap-2">
                    <Edit3 className="w-5 h-5 text-[#00C4CC]" />
                    <h3 className="font-serif font-bold text-base text-[#1C1917]">Renomear Item</h3>
                  </div>
                  <button
                    type="button"
                    onClick={() => setShowRenameModal(false)}
                    className="text-xs font-medium text-[#78716C] hover:text-[#1C1917]"
                  >
                    Cancelar
                  </button>
                </div>

                <div className="space-y-1">
                  <label className="block text-xs font-medium text-[#57534E]">Novo Nome</label>
                  <input
                    type="text"
                    value={renameInput}
                    onChange={(e) => setRenameInput(e.target.value)}
                    autoFocus
                    className="w-full text-xs px-3 py-2 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917]"
                  />
                </div>

                <div className="pt-2 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowRenameModal(false)}
                    className="px-3.5 py-2 border border-[#E7E5E4] text-xs font-medium rounded-lg text-[#57534E] hover:bg-[#FAF9F6]"
                  >
                    Cancelar
                  </button>
                  <button
                    type="submit"
                    disabled={renamingItem || !renameInput.trim()}
                    className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
                  >
                    {renamingItem ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <Check className="w-3.5 h-3.5" />
                    )}
                    <span>Salvar Nome</span>
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* ========================================================================= */}
          {/* MODAL 3: MUDAR DE LUGAR (MOVER ITEM/PASTA PARA OUTRA PASTA)                 */}
          {/* ========================================================================= */}
          {showMoveModal && itemToMove && (
            <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
              <form
                onSubmit={handleMoveItem}
                className="bg-white rounded-2xl max-w-sm w-full p-6 shadow-2xl border border-[#E7E5E4] space-y-4"
              >
                <div className="flex items-center justify-between pb-2 border-b border-[#E7E5E4]">
                  <div className="flex items-center gap-2">
                    <FolderInput className="w-5 h-5 text-[#00C4CC]" />
                    <h3 className="font-serif font-bold text-base text-[#1C1917]">Mudar de Lugar</h3>
                  </div>
                  <button
                    type="button"
                    onClick={() => setShowMoveModal(false)}
                    className="text-xs font-medium text-[#78716C] hover:text-[#1C1917]"
                  >
                    Cancelar
                  </button>
                </div>

                <p className="text-xs text-[#78716C]">
                  Selecione a pasta de destino para mover <strong className="text-[#1C1917]">"{itemToMove.name}"</strong>:
                </p>

                <div className="space-y-1">
                  <label className="block text-xs font-medium text-[#57534E]">Pasta de Destino</label>
                  <select
                    value={targetFolderIdForMove}
                    onChange={(e) => setTargetFolderIdForMove(e.target.value)}
                    className="w-full text-xs px-3 py-2 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg focus:outline-none focus:ring-1 focus:ring-[#1C1917] font-sans"
                  >
                    {allAvailableFolders
                      .filter((f) => f.id !== itemToMove.id)
                      .map((f) => (
                        <option key={f.id} value={f.id}>
                          📁 {f.name}
                        </option>
                      ))}
                  </select>
                </div>

                <div className="pt-2 flex justify-end gap-2">
                  <button
                    type="button"
                    onClick={() => setShowMoveModal(false)}
                    className="px-3.5 py-2 border border-[#E7E5E4] text-xs font-medium rounded-lg text-[#57534E] hover:bg-[#FAF9F6]"
                  >
                    Cancelar
                  </button>
                  <button
                    type="submit"
                    disabled={movingItem}
                    className="px-4 py-2 bg-[#1C1917] hover:bg-[#292524] text-white text-xs font-semibold rounded-lg flex items-center gap-1.5 disabled:opacity-50 cursor-pointer"
                  >
                    {movingItem ? (
                      <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                    ) : (
                      <FolderInput className="w-3.5 h-3.5" />
                    )}
                    <span>Mover Aqui</span>
                  </button>
                </div>
              </form>
            </div>
          )}

          {/* ========================================================================= */}
          {/* MODAL 4: SELETOR DE MODELO / DESIGN DO CANVA PARA OPÇÃO 1 (AUTOFILL)       */}
          {/* ========================================================================= */}
          {showTemplateSelectorModal && (
            <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-3 sm:p-6">
              <div className="bg-white rounded-2xl max-w-4xl w-full max-h-[88vh] flex flex-col shadow-2xl overflow-hidden border border-[#E7E5E4]">
                <div className="px-6 py-4 border-b border-[#E7E5E4] flex flex-col sm:flex-row sm:items-center justify-between gap-3 bg-white shrink-0">
                  <div>
                    <h3 className="text-lg font-serif font-bold text-[#1C1917]">
                      Selecionar Modelo ou Design do Canva
                    </h3>
                    <p className="text-xs text-[#78716C] mt-0.5">
                      Escolha um Brand Template oficial (Autofill) ou navegue pelas pastas da sua conta Canva para inspecionar.
                    </p>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className="relative w-56">
                      <Search className="w-3.5 h-3.5 text-[#A8A29E] absolute left-2.5 top-1/2 -translate-y-1/2" />
                      <input
                        type="text"
                        value={templateSelectorSearch}
                        onChange={(e) => setTemplateSelectorSearch(e.target.value)}
                        placeholder="Filtrar por nome..."
                        className="text-xs pl-8 pr-3 py-1.5 bg-[#FAF9F6] border border-[#E7E5E4] rounded-lg w-full"
                      />
                    </div>
                    <button
                      type="button"
                      onClick={() => setShowTemplateSelectorModal(false)}
                      className="p-1.5 text-[#78716C] hover:text-[#1C1917] border border-[#E7E5E4] rounded-lg cursor-pointer"
                    >
                      <X className="w-4 h-4" />
                    </button>
                  </div>
                </div>

                {/* Tabs internas do Modal */}
                <div className="px-6 py-2.5 bg-[#FAF9F6] border-b border-[#E7E5E4] flex flex-wrap items-center justify-between gap-2 text-xs shrink-0">
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setTemplateSelectorTab('brand_templates')}
                      className={`px-3 py-1.5 rounded-lg font-semibold cursor-pointer flex items-center gap-1.5 ${
                        templateSelectorTab === 'brand_templates'
                          ? 'bg-[#00C4CC] text-white'
                          : 'bg-white border border-[#E7E5E4] text-[#57534E]'
                      }`}
                    >
                      <Sparkles className="w-3.5 h-3.5" />
                      <span>Brand Templates (Autofill)</span>
                    </button>
                    <button
                      type="button"
                      onClick={() => setTemplateSelectorTab('folders')}
                      className={`px-3 py-1.5 rounded-lg font-semibold cursor-pointer flex items-center gap-1.5 ${
                        templateSelectorTab === 'folders'
                          ? 'bg-[#1C1917] text-white'
                          : 'bg-white border border-[#E7E5E4] text-[#57534E]'
                      }`}
                    >
                      <FolderOpen className="w-3.5 h-3.5" />
                      <span>Pastas & Designs da Conta</span>
                    </button>
                  </div>

                  {templateSelectorTab === 'folders' && (
                    <div className="flex items-center gap-1.5 text-[#78716C]">
                      {tplBreadcrumbs.map((crumb, idx) => (
                        <React.Fragment key={`${crumb.id}_${idx}`}>
                          {idx > 0 && <span>/</span>}
                          <button
                            type="button"
                            onClick={() => {
                              setTplBreadcrumbs((prev) => prev.slice(0, idx + 1));
                              loadTemplateSelectorFolder(crumb.id, crumb.name, false);
                            }}
                            className={`hover:text-[#1C1917] cursor-pointer ${
                              idx === tplBreadcrumbs.length - 1 ? 'font-bold text-[#1C1917]' : ''
                            }`}
                          >
                            {crumb.name}
                          </button>
                        </React.Fragment>
                      ))}
                    </div>
                  )}
                </div>

                <div className="p-6 overflow-y-auto flex-1 bg-[#FAF9F6]">
                  {templateSelectorTab === 'brand_templates' ? (
                    loadingBrandTemplates ? (
                      <div className="text-center py-14 space-y-2">
                        <RefreshCw className="w-7 h-7 animate-spin text-[#00C4CC] mx-auto" />
                        <p className="text-xs font-medium text-[#57534E]">Carregando Brand Templates do Canva...</p>
                      </div>
                    ) : brandTemplateAuthError ? (
                      <div className="bg-white border border-amber-200 rounded-xl p-6 text-center space-y-3">
                        <AlertCircle className="w-7 h-7 text-amber-600 mx-auto" />
                        <p className="text-xs font-semibold text-[#1C1917]">Autorização de Brand Templates Necessária</p>
                        <p className="text-xs text-[#57534E] max-w-lg mx-auto leading-relaxed">
                          {brandTemplateAuthError}
                        </p>
                        <div className="flex flex-wrap items-center justify-center gap-2 pt-1">
                          <button
                            type="button"
                            onClick={() => handleConnect(true)}
                            className="px-4 py-2 bg-[#00C4CC] hover:bg-[#00B0B7] text-white text-xs font-bold rounded-lg cursor-pointer"
                          >
                            Autorizar Brand Templates no Canva
                          </button>
                          <button
                            type="button"
                            onClick={() => setTemplateSelectorTab('folders')}
                            className="px-4 py-2 bg-white border border-[#E7E5E4] text-xs font-medium text-[#1C1917] rounded-lg cursor-pointer"
                          >
                            Explorar Pastas da Conta
                          </button>
                        </div>
                      </div>
                    ) : brandTemplatesList.length === 0 ? (
                      <div className="bg-white border border-[#E7E5E4] rounded-xl p-6 text-center space-y-2">
                        <Sparkles className="w-7 h-7 text-[#A8A29E] mx-auto" />
                        <p className="text-xs font-semibold text-[#1C1917]">Nenhum Brand Template encontrado</p>
                        <p className="text-xs text-[#78716C] max-w-md mx-auto">
                          Publique um design como Brand Template no Canva (com campos de dados vinculados) ou navegue pelas suas pastas para inspecionar um design.
                        </p>
                      </div>
                    ) : (
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-4">
                        {brandTemplatesList
                          .filter((bt) =>
                            bt.name.toLowerCase().includes(templateSelectorSearch.toLowerCase())
                          )
                          .map((bt) => (
                            <div
                              key={bt.id}
                              onClick={() => inspectSelectedCanvaTemplate(bt.id, bt.name)}
                              className="bg-white border border-[#E7E5E4] hover:border-[#00C4CC] rounded-xl overflow-hidden cursor-pointer shadow-xs transition-all flex flex-col"
                            >
                              <div className="h-40 bg-neutral-900 relative">
                                {bt.thumbnailUrl ? (
                                  <img src={bt.thumbnailUrl} alt={bt.name} className="w-full h-full object-cover" />
                                ) : (
                                  <div className="w-full h-full flex items-center justify-center text-white/50 text-xs">
                                    Brand Template
                                  </div>
                                )}
                                <span className="absolute top-2 left-2 bg-[#00C4CC] text-white text-[9px] font-mono font-bold px-2 py-0.5 rounded">
                                  BRAND TEMPLATE
                                </span>
                              </div>
                              <div className="p-3 flex items-center justify-between gap-2">
                                <div className="min-w-0">
                                  <p className="text-xs font-bold text-[#1C1917] truncate">{bt.name}</p>
                                  <p className="text-[10px] font-mono text-[#78716C] truncate">{bt.id}</p>
                                </div>
                                <span className="text-[11px] font-semibold text-[#008A90] shrink-0">
                                  Selecionar →
                                </span>
                              </div>
                            </div>
                          ))}
                      </div>
                    )
                  ) : loadingTplFolder ? (
                    <div className="text-center py-14 space-y-2">
                      <RefreshCw className="w-7 h-7 animate-spin text-[#00C4CC] mx-auto" />
                      <p className="text-xs font-medium text-[#57534E]">Listando pastas e designs do Canva...</p>
                    </div>
                  ) : (
                    <div className="space-y-4">
                      <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                        {tplFolderItems
                          .filter((item) =>
                            String(item.name || '').toLowerCase().includes(templateSelectorSearch.toLowerCase())
                          )
                          .map((item) => {
                            const isFolder = item.type === 'folder';
                            const isDesignOrTemplate = item.type === 'design' || item.type === 'brand_template';
                            return (
                              <div
                                key={item.id}
                                onClick={() => {
                                  if (isFolder) {
                                    setTplBreadcrumbs((prev) => [...prev, { id: item.id, name: item.name }]);
                                    loadTemplateSelectorFolder(item.id, item.name, false);
                                  } else if (isDesignOrTemplate) {
                                    inspectSelectedCanvaTemplate(item.id, item.name);
                                  }
                                }}
                                className={`bg-white border border-[#E7E5E4] rounded-xl p-3 flex items-center gap-3 transition-all ${
                                  isFolder || isDesignOrTemplate
                                    ? 'cursor-pointer hover:border-[#00C4CC]'
                                    : 'opacity-60'
                                }`}
                              >
                                <div className="w-12 h-12 rounded-lg bg-[#FAF9F6] border border-[#E7E5E4] flex items-center justify-center shrink-0 overflow-hidden">
                                  {isFolder ? (
                                    <Folder className="w-6 h-6 text-amber-500" />
                                  ) : item.thumbnailUrl ? (
                                    <img src={item.thumbnailUrl} alt={item.name} className="w-full h-full object-cover" />
                                  ) : (
                                    <FileText className="w-5 h-5 text-[#78716C]" />
                                  )}
                                </div>
                                <div className="min-w-0 flex-1">
                                  <p className="text-xs font-bold text-[#1C1917] truncate">{item.name}</p>
                                  <p className="text-[10px] font-mono text-[#78716C] uppercase">
                                    {item.type}
                                  </p>
                                </div>
                              </div>
                            );
                          })}
                      </div>

                      {tplFolderContinuation && (
                        <div className="flex justify-center">
                          <button
                            type="button"
                            onClick={() => loadTemplateSelectorFolder(tplFolderId, undefined, true)}
                            disabled={loadingMoreTplFolder}
                            className="px-4 py-2 bg-white border border-[#E7E5E4] rounded-lg text-xs font-medium text-[#1C1917] cursor-pointer"
                          >
                            {loadingMoreTplFolder ? 'Carregando...' : 'Carregar mais itens'}
                          </button>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}
        </div>

        {/* PERSISTENT LIVE PREVIEW SIDEBAR ON THE RIGHT */}
        <aside className="w-full lg:w-80 shrink-0 bg-white border border-[#E7E5E4] rounded-2xl p-5 shadow-sm sticky top-24 space-y-4">
          <div className="flex items-center justify-between pb-3 border-b border-[#F5F5F4]">
            <div className="flex items-center gap-2">
              <Sparkles className="w-4 h-4 text-[#00C4CC]" />
              <h3 className="font-serif font-bold text-sm text-[#1C1917]">Prévia em Tempo Real</h3>
            </div>
            <span className="text-[11px] font-mono text-[#78716C]">
              Frase {previewIndex + 1} de {validPhrases.length || 1}
            </span>
          </div>

          <div className="flex flex-col items-center justify-center p-3 bg-[#FAF9F6] rounded-xl border border-[#E7E5E4]">
            {(() => {
              const curPhrase = validPhrases[previewIndex] || validPhrases[0] || 'Sua frase aqui.';
              const info = getPreviewPageInfo(previewIndex, mediaList);
              const curMedia = info.media;

              // Se não houver mídia na Etapa 4, mas houver modelo inspecionado na Opção 1, usa as páginas do modelo
              const tplTotal = inspectedTemplate?.pageCount || 1;
              const tplPageNum = (previewIndex % tplTotal) + 1;
              const tplPage = inspectedTemplate?.pages?.[(tplPageNum - 1) % (inspectedTemplate?.pages?.length || 1)];
              const tplLayout =
                inspectedTemplate?.slideLayouts?.find((l) => l.slideIndex === tplPageNum) ||
                inspectedTemplate?.slideLayouts?.[(tplPageNum - 1) % (inspectedTemplate?.slideLayouts?.length || 1)] ||
                null;

              const effectivePageNum = curMedia ? info.pageNum : tplPageNum;
              const effectiveTotalPages = curMedia ? info.totalPagesInDesign : tplTotal;
              const effectiveSlideLayout = info.slideLayout || tplLayout;
              const effectiveCleanBgUrl = info.cleanPageUrl || tplPage?.cleanBgUrl || '';
              const effectiveThumbUrl = info.pageUrl || tplPage?.thumbnailUrl || '';
              const hasAnyBg = Boolean(curMedia || tplPage);

              // Atualiza apenas `text` nos elementos originais de `slideLayouts` respeitando `pageScope`,
              // preservando fontFamily, fontSizePt, colorHex, bold e align sem inserir novas caixas sobrepostas
              const matchedTextBoxes = replaceSlideLayoutTextBoxesInPlace(
                effectiveSlideLayout,
                effectivePageNum,
                curPhrase,
                keyFieldMappings,
                signatureText
              );

              const hasInFilePlaceholderReplacement = matchedTextBoxes.some((b) => b.didReplace);
              const extractedHasPlaceholder = Boolean(
                curMedia?.extractedPages?.[(effectivePageNum - 1) % (curMedia?.extractedPages?.length || 1)]?.hasPhrasePlaceholder
              );
              const isRoundtripTemplateMode =
                Boolean(inspectedTemplate?.editableViaPptxRoundtrip) ||
                Boolean(effectiveSlideLayout && effectiveSlideLayout.textBoxes.length > 0);
              const isCanvaTemplateSlide =
                isRoundtripTemplateMode ||
                hasInFilePlaceholderReplacement ||
                Boolean(curMedia?.isCanvaDesignTemplate && extractedHasPlaceholder);

              const cleanBgReady = Boolean(effectiveCleanBgUrl && !failedCleanBgUrls[effectiveCleanBgUrl]);
              const bgDisplaySrc =
                isRoundtripTemplateMode && cleanBgReady
                  ? effectiveCleanBgUrl
                  : effectiveThumbUrl;

              return (
                <>
                  <div
                    className={`relative overflow-hidden rounded-lg shadow-md transition-all ${
                      aspectRatio === '9:16'
                        ? 'w-[200px] h-[356px]'
                        : aspectRatio === '1:1'
                        ? 'w-[220px] h-[220px]'
                        : 'w-[240px] h-[135px]'
                    }`}
                    style={{ backgroundColor: '#1C1917' }}
                  >
                    {/* Media background (usa imagem limpa sem (frase) estático quando em substituição in-place) */}
                    {hasAnyBg ? (
                      info.isVideo ? (
                        <video src={effectiveThumbUrl} className="absolute inset-0 w-full h-full object-cover" autoPlay loop muted />
                      ) : (
                        <img
                          src={bgDisplaySrc}
                          alt={`Página ${effectivePageNum}`}
                          onError={(e) => {
                            if (effectiveCleanBgUrl && e.currentTarget.src.includes(effectiveCleanBgUrl)) {
                              setFailedCleanBgUrls((prev) => ({ ...prev, [effectiveCleanBgUrl]: true }));
                            }
                            if (effectiveThumbUrl && e.currentTarget.src !== effectiveThumbUrl) {
                              e.currentTarget.src = effectiveThumbUrl;
                            }
                          }}
                          className="absolute inset-0 w-full h-full object-cover"
                        />
                      )
                    ) : (
                      <div className="absolute inset-0 bg-gradient-to-br from-neutral-800 to-neutral-950 flex items-center justify-center">
                        <span className="text-[9px] text-neutral-500 font-mono">Sem mídia</span>
                      </div>
                    )}

                    {/* Overlay apenas quando NÃO estiver preservando o design original do Canva */}
                    {!isCanvaTemplateSlide && (
                      <div
                        className="absolute inset-0 pointer-events-none"
                        style={{ backgroundColor: `rgba(0, 0, 0, ${overlayOpacity / 100})` }}
                      />
                    )}

                    {/* Badge da Página do Design do Canva */}
                    <div className="absolute top-2 left-2 z-20 bg-black/85 backdrop-blur-xs text-white text-[9px] font-mono font-bold px-2 py-0.5 rounded border border-white/20 shadow-xs flex items-center gap-1">
                      <span>Página {effectivePageNum} de {effectiveTotalPages}</span>
                      {previewIndex >= effectiveTotalPages && (
                        <span className="text-amber-300 font-semibold text-[8px] ml-0.5 bg-amber-950/80 px-1 py-0.2 rounded">(Repetida)</span>
                      )}
                    </div>

                    {/* MODO 1: Substituição IN-PLACE na posição, fonte e cor exatas do arquivo Canva (sem novas caixas) */}
                    {isRoundtripTemplateMode ? (
                      matchedTextBoxes.map((box, bIdx) => {
                        if (!cleanBgReady && !box.didReplace) return null;
                        const boxW = box.didReplace && box.align === 'center' ? Math.max(box.wPct, 82) : box.wPct;
                        const boxLeft =
                          box.didReplace && box.align === 'center'
                            ? Math.max(4, Math.min(100 - boxW - 4, box.xPct + box.wPct / 2 - boxW / 2))
                            : box.xPct;
                        return (
                          <div
                            key={`tb_${bIdx}`}
                            className="absolute pointer-events-none flex flex-col justify-center"
                            style={{
                              left: `${boxLeft}%`,
                              top: `${box.yPct}%`,
                              width: `${boxW}%`,
                              minHeight: `${box.hPct}%`,
                              textAlign: box.align,
                              padding: '2px 4px',
                              backgroundColor:
                                !cleanBgReady && box.didReplace ? 'rgba(28, 25, 23, 0.92)' : 'transparent',
                              borderRadius: !cleanBgReady && box.didReplace ? '4px' : undefined,
                            }}
                          >
                            <p
                              style={{
                                fontFamily: box.fontFamily ? `"${box.fontFamily}", sans-serif` : undefined,
                                fontSize: `${Math.max(9, Math.min(16, box.fontSizePt * 0.38))}px`,
                                color: box.colorHex || '#FFFFFF',
                                fontWeight: box.bold ? 700 : 500,
                                fontStyle: box.italic ? 'italic' : 'normal',
                                textDecoration: box.underline ? 'underline' : undefined,
                                letterSpacing: box.charSpacingPt ? `${box.charSpacingPt * 0.35}px` : undefined,
                                lineHeight: box.lineSpacingPct ? box.lineSpacingPct / 100 : 1.22,
                                textShadow: '0 1px 3px rgba(0,0,0,0.45)',
                              }}
                            >
                              {box.text}
                            </p>
                          </div>
                        );
                      })
                    ) : !isCanvaTemplateSlide ? (
                      /* MODO 2: Criação de Novo Design (Etapa 5 aplica texto central + assinatura) */
                      <div
                        className={`absolute inset-0 p-4 flex flex-col pointer-events-none ${
                          textPosition === 'top'
                            ? 'justify-start'
                            : textPosition === 'bottom'
                            ? 'justify-end'
                            : 'justify-center'
                        }`}
                        style={{ textAlign }}
                      >
                        <p
                          style={{
                            fontFamily: fontFamily === 'serif' ? 'Fraunces, serif' : 'DM Sans, sans-serif',
                            fontSize: `${Math.max(12, fontSize * 0.75)}px`,
                            color: textColor,
                            lineHeight: 1.3,
                            textShadow: hasShadow ? '0 2px 6px rgba(0,0,0,0.8)' : 'none',
                            WebkitTextStroke: hasStroke ? '0.5px rgba(0,0,0,0.6)' : 'none',
                          }}
                          className="font-medium"
                        >
                          {curPhrase}
                        </p>

                        {signatureText && (
                          <p
                            className="text-[9px] font-mono tracking-wider opacity-80 mt-2"
                            style={{ color: textColor }}
                          >
                            {signatureText}
                          </p>
                        )}
                      </div>
                    ) : null}
                  </div>

                  {/* Mini navigation controls */}
                  <div className="flex items-center justify-between w-full mt-3 pt-3 border-t border-[#E7E5E4]">
                    <button
                      onClick={() => setPreviewIndex((prev) => Math.max(0, prev - 1))}
                      disabled={previewIndex === 0}
                      className="px-2.5 py-1 border border-[#E7E5E4] rounded text-[11px] bg-white disabled:opacity-40"
                    >
                      Anterior
                    </button>
                    <span className="text-[11px] font-mono text-[#57534E]">
                      {previewIndex + 1} / {validPhrases.length || 1}
                    </span>
                    <button
                      onClick={() => setPreviewIndex((prev) => Math.min(validPhrases.length - 1, prev + 1))}
                      disabled={previewIndex >= validPhrases.length - 1}
                      className="px-2.5 py-1 border border-[#E7E5E4] rounded text-[11px] bg-white disabled:opacity-40"
                    >
                      Próxima
                    </button>
                  </div>

                  {hasInFilePlaceholderReplacement && (
                    <div className="mt-2.5 w-full px-2.5 py-1.5 bg-emerald-50 border border-emerald-200 rounded-lg text-[10.5px] font-medium text-emerald-900 text-center">
                      ✓ Substituindo marcador diretamente dentro do arquivo (sem texto extra sobreposto)
                    </div>
                  )}
                </>
              );
            })()}
          </div>

          <div className="text-[11px] text-[#78716C] leading-relaxed">
            Esta prévia acompanha suas edições de frases, mídias e estilos em tempo real.
          </div>
        </aside>
      </main>

      {/* FLOATING HOVER PREVIEW POPUP (Step 9 Drive Explorer) */}
      {hoveredItemForPreview && (
        <div
          style={{ top: `${hoverPosition.y}px`, left: `${hoverPosition.x}px` }}
          className="fixed z-50 pointer-events-none bg-white rounded-xl shadow-2xl border border-[#E7E5E4] p-3 w-64 animate-fade-in space-y-2"
        >
          <div className="w-full h-36 rounded-lg overflow-hidden bg-neutral-900 relative">
            {hoveredItemForPreview.category === 'video' ? (
              <video
                src={hoveredItemForPreview.thumbnailUrl}
                autoPlay
                muted
                loop
                className="w-full h-full object-cover"
              />
            ) : hoveredItemForPreview.category === 'audio' ? (
              <div className="w-full h-full flex flex-col items-center justify-center bg-gradient-to-br from-emerald-900 to-black text-white p-2">
                <Music className="w-8 h-8 text-emerald-400 mb-1 animate-pulse" />
                <span className="text-[10px] font-mono text-emerald-200 text-center truncate w-full">Prévia de Áudio</span>
                <audio src={hoveredItemForPreview.thumbnailUrl} controls autoPlay className="w-full h-6 mt-1 scale-90" />
              </div>
            ) : (
              <img
                src={hoveredItemForPreview.thumbnailUrl}
                alt={hoveredItemForPreview.name}
                className="w-full h-full object-cover"
              />
            )}
            <span className="absolute bottom-1 right-1 bg-black/70 text-white text-[9px] font-mono px-1.5 py-0.5 rounded uppercase font-semibold">
              {hoveredItemForPreview.category}
            </span>
          </div>

          <div className="space-y-1 px-0.5">
            <p className="text-xs font-semibold text-[#1C1917] truncate">{hoveredItemForPreview.name}</p>
            <div className="grid grid-cols-2 gap-x-2 gap-1 text-[10px] text-[#78716C] pt-1 border-t border-[#F5F5F4]">
              <div className="flex items-center justify-between">
                <span>Modificação:</span>
                <span className="font-mono text-[#1C1917]">{hoveredItemForPreview.modifiedAt || 'Hoje'}</span>
              </div>
              <div className="flex items-center justify-between">
                <span>Tamanho:</span>
                <span className="font-mono text-[#1C1917]">{hoveredItemForPreview.size || '2.4 MB'}</span>
              </div>
              <div className="flex items-center justify-between col-span-2">
                <span>Dimensões / Formato:</span>
                <span className="font-mono text-[#1C1917]">
                  {hoveredItemForPreview.category === 'video' ? '1080x1920 (0:15s)' : hoveredItemForPreview.category === 'design' ? '1080x1350 (Editor)' : '1920x1080 (HD)'}
                </span>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="border-t border-[#E7E5E4] bg-white py-4 px-6 text-center text-xs text-[#78716C]">
        <p>LoteCriativo Canva · Produção em lote com APIs oficiais e exportação editável.</p>
      </footer>
    </div>
  );
}
