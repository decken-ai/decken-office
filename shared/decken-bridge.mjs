export const DEFAULT_API_ORIGIN = 'https://api.decken.ai';

function isLoopback(hostname) {
  const h = hostname.toLowerCase();
  return h === 'localhost' || h === '127.0.0.1' || h === '[::1]' || h.endsWith('.localhost');
}

export function validateApiOrigin(origin = DEFAULT_API_ORIGIN, hasApiKey = false) {
  const parsed = new URL(origin.trim());
  if (parsed.protocol !== 'https:') {
    if (parsed.protocol === 'http:' && isLoopback(parsed.hostname)) {
      return parsed.origin;
    }
    if (hasApiKey || parsed.protocol !== 'http:') {
      throw new Error(`Refusing non-HTTPS API origin (${origin}) when API key is configured.`);
    }
  }
  return parsed.origin;
}

export async function generateDeckViaDecken({
  outline,
  reference,
  source,
  prompt,
  deckTitle,
  engineMode = 'vlm',
  modelId,
  apiKey = '',
  apiOrigin = DEFAULT_API_ORIGIN,
  fetchImpl = globalThis.fetch,
}) {
  if (!source && !outline && !prompt) {
    throw new Error('`generateDeckViaDecken` requires `outline` / `prompt` or `source`.');
  }
  const cleanKey = apiKey.trim();
  const base = validateApiOrigin(apiOrigin, Boolean(cleanKey));
  const headers = { 'Content-Type': 'application/json' };
  if (cleanKey) {
    headers.Authorization = `Bearer ${cleanKey}`;
  }

  const res = await fetchImpl(`${base}/v1/decks`, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      mode: 'sync',
      engineMode,
      ...(deckTitle ? { deckTitle } : {}),
      ...(prompt ? { prompt } : {}),
      ...(outline !== undefined ? { outline } : {}),
      ...(reference !== undefined ? { reference } : {}),
      ...(source !== undefined ? { source } : {}),
      ...(modelId ? { modelId } : {}),
      includePptxBase64: true,
    }),
  });

  const data = await res.json();
  if (!res.ok || data.status === 'failed') {
    throw new Error(data.error || `Decken generation failed (${res.status})`);
  }
  return data;
}

export async function reconstructSlideViaDecken({
  imageDataUrl,
  slideId = 'office-addin-slide',
  engineMode = 'vlm',
  modelId,
  apiKey = '',
  apiOrigin = DEFAULT_API_ORIGIN,
  fetchImpl = globalThis.fetch,
}) {
  if (!imageDataUrl || typeof imageDataUrl !== 'string') {
    throw new Error('`imageDataUrl` is required.');
  }
  const cleanKey = apiKey.trim();
  const base = validateApiOrigin(apiOrigin, Boolean(cleanKey));
  const headers = { 'Content-Type': 'application/json' };
  if (cleanKey) {
    headers.Authorization = `Bearer ${cleanKey}`;
  }

  const res = await fetchImpl(`${base}/v1/reconstruct`, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      slideId,
      imageDataUrl,
      mode: 'sync',
      engineMode,
      ...(modelId ? { modelId } : {}),
      includePptxBase64: true,
    }),
  });

  const data = await res.json();
  if (!res.ok || data.status === 'failed') {
    throw new Error(data.error || `Decken reconstruction failed (${res.status})`);
  }
  return data;
}

/**
 * Inserts a Base64 .pptx slide into an active Microsoft PowerPoint presentation via Office.js 1.5+.
 */
export async function insertIntoMicrosoftPowerPoint(pptxBase64, options = {}) {
  const host = options.powerPointGlobal ?? globalThis.PowerPoint;
  if (!host?.run) {
    throw new Error('Microsoft PowerPoint Office.js runtime (`PowerPoint.run`) is not available.');
  }
  const formattingMode =
    options.formatting === 'theme'
      ? (host.InsertSlideFormatting?.useDestinationTheme ?? 'UseDestinationTheme')
      : (host.InsertSlideFormatting?.keepSourceFormatting ?? 'KeepSourceFormatting');

  await host.run(async (context) => {
    let targetSlideId = options.targetSlideId;
    if (!targetSlideId && context.presentation.getSelectedSlides) {
      const selected = context.presentation.getSelectedSlides();
      selected.load('items/id');
      await context.sync();
      targetSlideId = selected.items?.[0]?.id;
    }
    context.presentation.insertSlidesFromBase64(pptxBase64, {
      formatting: formattingMode,
      ...(targetSlideId ? { targetSlideId } : {}),
    });
    await context.sync();
  });
}

/**
 * Inserts a Base64 .pptx slide into an active WPS Office Presentation (WPS演示 / wpp) via WPS JSAPI.
 */
export function insertIntoWpsPresentation(pptxBase64, options = {}) {
  const wpsGlobal = options.wpsGlobal ?? globalThis.wps;
  if (!wpsGlobal?.WpsApplication || !wpsGlobal?.FileSystem || !wpsGlobal?.Env) {
    throw new Error('WPS Presentation JSAPI runtime (`wps.WpsApplication`) is not available.');
  }

  const app = wpsGlobal.WpsApplication();
  const presentation = app.ActivePresentation;
  if (!presentation) {
    throw new Error('No active WPS presentation open.');
  }

  const tempDir = wpsGlobal.Env.GetTempPath().replace(/[/\\]+$/, '');
  const tempFile = `${tempDir}/decken_reconstruct_${Date.now()}.pptx`;
  const binaryString = typeof atob === 'function'
    ? atob(pptxBase64)
    : Buffer.from(pptxBase64, 'base64').toString('binary');

  wpsGlobal.FileSystem.writeAsBinaryString(tempFile, binaryString);

  let afterIndex = presentation.Slides?.Count ?? 0;
  try {
    const selSlide = app.ActiveWindow?.Selection?.SlideRange?.Item?.(1);
    if (selSlide?.SlideIndex) {
      afterIndex = selSlide.SlideIndex;
    }
  } catch {
    // Fallback to appending at end of presentation
  }

  const insertedCount = presentation.Slides.InsertFromFile(tempFile, afterIndex);
  return { tempFile, afterIndex, insertedCount };
}
