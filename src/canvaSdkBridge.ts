/**
 * Canva Apps SDK Runtime Bridge & Capabilities Detector
 *
 * Fornece integração compatível com o Canva Apps SDK (@canva/design) quando o app
 * é executado embarcado dentro do Editor do Canva (iframe de Canva App):
 * - Selection API (selection.read para 'video' | 'audio' | 'richtext')
 * - getTemporaryUrl() com transferência imediata para o backend do app
 * - Richtext Editing API: editContent, readTextRegions, readPlaintext (somente leitura)
 *   e replaceText(range, newText) preservando formatação por região (sem sobrescrever com Plaintext).
 *
 * Quando executado fora do Editor do Canva (aplicativo web externo conectado via Canva Connect OAuth),
 * detecta o ambiente com transparência e NUNCA simula dados fictícios.
 */

export interface CanvaTextRegion {
  text: string;
  formatting?: {
    fontFamily?: string;
    fontSize?: number;
    color?: string;
    fontWeight?: string;
    fontStyle?: string;
    textAlign?: string;
    letterSpacing?: number;
    lineHeight?: number;
    decoration?: string;
  };
}

export interface CanvaRichtextSurface {
  readTextRegions?: () => Promise<CanvaTextRegion[]>;
  readPlaintext?: () => Promise<string>;
  replaceText?: (
    range: { index: number; length: number },
    text: string,
    formatting?: Record<string, any>
  ) => Promise<void> | void;
}

export interface CanvaSelectedMediaElement {
  id?: string;
  type: 'video' | 'audio' | string;
  name?: string;
  mimeType?: string;
  durationInSeconds?: number;
  getTemporaryUrl?: () => Promise<string | { url: string }>;
  ref?: string;
}

export interface CanvaSdkHostStatus {
  insideCanvaEditor: boolean;
  hasSelectionApi: boolean;
  hasRichtextEditContent: boolean;
  hasGetTemporaryUrl: boolean;
  modeLabel: string;
  explanation: string;
}

export interface TransferredSdkMediaResult {
  success: boolean;
  transferred: Array<{
    id: string;
    name: string;
    type: 'video' | 'audio';
    mimeType: string;
    size: number;
    url: string;
    source: 'canva_apps_sdk_selection';
  }>;
  error?: string;
}

export interface RichtextReplacementDetail {
  elementIndex: number;
  matchedKeyword: string;
  replacedWith: string;
  startIndex: number;
  length: number;
  preservedFormatting?: CanvaTextRegion['formatting'];
}

function getGlobalCanvaRuntime(): any {
  if (typeof window === 'undefined') return null;
  const w = window as any;
  // Detecta runtime oficial injetado pelo Canva App SDK host ou bridge registrado em window
  return w.__CANVA_APPS_SDK__ || w.canvaDesignSdk || w.canva?.design || null;
}

export function detectCanvaAppsSdkHost(): CanvaSdkHostStatus {
  const sdk = getGlobalCanvaRuntime();
  const isIframe = typeof window !== 'undefined' && window.self !== window.top;
  const referrerIsCanva =
    typeof document !== 'undefined' &&
    /canva\.(com|cn)/i.test(document.referrer || '');

  const hasSelection = Boolean(sdk?.selection?.read);
  const hasRichtext = Boolean(sdk?.selection?.read || sdk?.editContent);
  const insideCanvaEditor = Boolean(sdk && (hasSelection || hasRichtext));

  if (insideCanvaEditor) {
    return {
      insideCanvaEditor: true,
      hasSelectionApi: hasSelection,
      hasRichtextEditContent: hasRichtext,
      hasGetTemporaryUrl: true,
      modeLabel: 'Canva Apps SDK Ativo (Editor do Canva)',
      explanation:
        'Selection API, getTemporaryUrl e Richtext (editContent / readTextRegions / replaceText) disponíveis diretamente no design aberto.',
    };
  }

  return {
    insideCanvaEditor: false,
    hasSelectionApi: false,
    hasRichtextEditContent: false,
    hasGetTemporaryUrl: false,
    modeLabel:
      isIframe && referrerIsCanva
        ? 'Iframe Canva sem SDK @canva/design carregado'
        : 'Aplicativo Web Externo (Canva Connect API OAuth)',
    explanation:
      'Selection API, getTemporaryUrl e editContent/Richtext pertencem ao Canva Apps SDK (executável apenas dentro do Editor do Canva). No modo Web Externo, usamos a Canva Connect REST API (Folders, Designs, Assets ID, Brand Template Autofill e Round-Trip PPTX XML).',
  };
}

/**
 * Lê vídeos ou áudios selecionados no Editor do Canva via Selection API,
 * chama getTemporaryUrl() imediatamente e envia os bytes para /api/canva/media/transfer-temp-url
 * antes que a URL temporária expire.
 */
export async function captureSelectedMediaFromCanvaEditor(
  mediaKind: 'video' | 'audio'
): Promise<TransferredSdkMediaResult> {
  const sdk = getGlobalCanvaRuntime();
  if (!sdk?.selection?.read) {
    return {
      success: false,
      transferred: [],
      error:
        'Selection API do Canva Apps SDK indisponível fora do Editor do Canva. Para importar fora do editor, use Asset ID, Exportação MP4 de um design de vídeo ou Upload local.',
    };
  }

  try {
    const selectionScope = await sdk.selection.read(mediaKind);
    const items: CanvaSelectedMediaElement[] = Array.isArray(selectionScope?.contents)
      ? selectionScope.contents
      : Array.isArray(selectionScope)
      ? selectionScope
      : [];

    if (items.length === 0) {
      return {
        success: false,
        transferred: [],
        error: `Nenhum elemento de ${mediaKind === 'video' ? 'vídeo' : 'áudio'} selecionado no canvas do Canva.`,
      };
    }

    const transferred: TransferredSdkMediaResult['transferred'] = [];

    for (let i = 0; i < items.length; i++) {
      const el = items[i];
      if (typeof el.getTemporaryUrl !== 'function') {
        continue;
      }

      // Regra Canva Apps SDK: obter URL temporária e transferir imediatamente para armazenamento persistente
      const tempRes = await el.getTemporaryUrl();
      const tempUrl = typeof tempRes === 'string' ? tempRes : tempRes?.url;
      if (!tempUrl) continue;

      const transferRes = await fetch('/api/canva/media/transfer-temp-url', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          temporaryUrl: tempUrl,
          mediaType: mediaKind,
          name: el.name || `${mediaKind === 'video' ? 'Video_Canva' : 'Audio_Canva'}_${Date.now()}_${i + 1}`,
          mimeType: el.mimeType || (mediaKind === 'video' ? 'video/mp4' : 'audio/mpeg'),
          ref: el.ref || el.id,
        }),
      });

      const transferData = await transferRes.json();
      if (transferRes.ok && transferData.success && transferData.media) {
        transferred.push({
          ...transferData.media,
          source: 'canva_apps_sdk_selection',
        });
      }
    }

    if (transferred.length === 0) {
      return {
        success: false,
        transferred: [],
        error: 'Não foi possível transferir a URL temporária do elemento selecionado.',
      };
    }

    return { success: true, transferred };
  } catch (err: any) {
    return {
      success: false,
      transferred: [],
      error: err?.message || 'Erro ao ler seleção via Canva Apps SDK.',
    };
  }
}

/**
 * Executa substituição de palavras-chave via Richtext oficial do Canva Apps SDK:
 * - Lê regiões existentes via readTextRegions() (preservando offsets e estilos individuais de cada run)
 * - Usa readPlaintext() exclusivamente como leitura auxiliar caso readTextRegions não retorne texto
 * - Aplica replaceText({ index, length }, replacement) DENTRO de editContent, mantendo
 *   fonte, tamanho, cor, negrito, itálico, alinhamento, espaçamento e geometria da caixa.
 * - NUNCA sobrescreve o conteúdo via plaintext completo.
 */
export async function replaceKeywordsInCanvaRichtextSelection(
  replacements: Array<{ findText: string; replaceText: string }>
): Promise<{
  success: boolean;
  replacedCount: number;
  details: RichtextReplacementDetail[];
  error?: string;
}> {
  const sdk = getGlobalCanvaRuntime();
  if (!sdk?.selection?.read) {
    return {
      success: false,
      replacedCount: 0,
      details: [],
      error:
        'editContent / Richtext do Canva Apps SDK só executa quando o app está aberto dentro do Editor do Canva.',
    };
  }

  try {
    const selectionScope = await sdk.selection.read('richtext');
    const selectedRichtexts = Array.isArray(selectionScope?.contents)
      ? selectionScope.contents
      : selectionScope && (typeof selectionScope.editContent === 'function' || typeof selectionScope.replaceText === 'function')
      ? [selectionScope]
      : [];

    if (selectedRichtexts.length === 0) {
      return {
        success: false,
        replacedCount: 0,
        details: [],
        error: 'Selecione ao menos uma caixa de texto (Richtext) no Editor do Canva.',
      };
    }

    const details: RichtextReplacementDetail[] = [];

    for (let elIdx = 0; elIdx < selectedRichtexts.length; elIdx++) {
      const rtRef = selectedRichtexts[elIdx];
      const runEdit = async (surface: CanvaRichtextSurface) => {
        if (typeof surface.replaceText !== 'function') {
          throw new Error('Superfície Richtext sem suporte a replaceText.');
        }

        // 1. Inspeciona regiões formatadas preservando metadados visuais (sem destruir estilos)
        const regions: CanvaTextRegion[] =
          typeof surface.readTextRegions === 'function'
            ? await surface.readTextRegions()
            : [];

        // 2. readPlaintext somente para localizar offsets exatos caso readTextRegions esteja vazio
        let currentText =
          regions.length > 0
            ? regions.map((r) => r.text || '').join('')
            : typeof surface.readPlaintext === 'function'
            ? await surface.readPlaintext()
            : '';

        for (const rep of replacements) {
          const target = String(rep.findText || '').trim();
          if (!target || !currentText) continue;

          const escaped = target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const hasBracket = /^[\(\{\[<].*[\)\}\]>]$/.test(target);
          const candidatePatterns = hasBracket
            ? [new RegExp(escaped, 'gi')]
            : [
                new RegExp(`\\(${escaped}\\)`, 'gi'),
                new RegExp(`\\{${escaped}\\}`, 'gi'),
                new RegExp(`\\[${escaped}\\]`, 'gi'),
                new RegExp(`<${escaped}>`, 'gi'),
                new RegExp(escaped, 'gi'),
              ];

          for (const pat of candidatePatterns) {
            let matchFoundInPattern = false;
            let cursor = 0;
            while (cursor < currentText.length) {
              pat.lastIndex = 0;
              const m = pat.exec(currentText.slice(cursor));
              if (!m || m[0].length === 0) break;

              matchFoundInPattern = true;
              const matchIndex = cursor + m.index;
              const matchLen = m[0].length;

              // Identifica a formatação da região onde o marcador começa para log de auditoria
              let acc = 0;
              let matchedStyle: CanvaTextRegion['formatting'] | undefined;
              for (const reg of regions) {
                const regLen = (reg.text || '').length;
                if (matchIndex >= acc && matchIndex < acc + regLen) {
                  matchedStyle = reg.formatting;
                  break;
                }
                acc += regLen;
              }

              // Substituição cirúrgica via Richtext.replaceText preserving original region formatting
              await surface.replaceText(
                { index: matchIndex, length: matchLen },
                rep.replaceText
              );

              currentText =
                currentText.slice(0, matchIndex) +
                rep.replaceText +
                currentText.slice(matchIndex + matchLen);

              details.push({
                elementIndex: elIdx,
                matchedKeyword: m[0],
                replacedWith: rep.replaceText,
                startIndex: matchIndex,
                length: matchLen,
                preservedFormatting: matchedStyle,
              });

              cursor = matchIndex + rep.replaceText.length;
            }
            if (matchFoundInPattern) break;
          }
        }
      };

      if (typeof rtRef.editContent === 'function') {
        await rtRef.editContent(runEdit);
      } else if (typeof selectionScope.save === 'function') {
        await runEdit(rtRef);
        await selectionScope.save();
      } else {
        await runEdit(rtRef);
      }
    }

    return {
      success: details.length > 0,
      replacedCount: details.length,
      details,
      error:
        details.length === 0
          ? 'Nenhuma palavra-chave configurada foi encontrada nos textos Richtext selecionados.'
          : undefined,
    };
  } catch (err: any) {
    return {
      success: false,
      replacedCount: 0,
      details: [],
      error: err?.message || 'Falha ao editar Richtext no Canva Apps SDK.',
    };
  }
}
