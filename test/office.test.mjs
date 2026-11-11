import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import test from 'node:test';
import {
  generateDeckViaDecken,
  insertIntoMicrosoftPowerPoint,
  insertIntoWpsPresentation,
  reconstructSlideViaDecken,
  validateApiOrigin,
} from '../shared/decken-bridge.mjs';

test('PowerPoint manifest.xml and WPS ribbon.xml pass structural & branding checks, and host bridges insert native slides', async () => {
  // 1. Validate Microsoft PowerPoint manifest.xml
  const pptManifest = await readFile(new URL('../powerpoint/manifest.xml', import.meta.url), 'utf8');
  assert.match(pptManifest, /<ProviderName>Decken<\/ProviderName>/);
  assert.doesNotMatch(pptManifest, /LYU\.ai/i);
  assert.match(pptManifest, /<Set Name="PowerPointApi" MinVersion="1\.5"\/>/);
  assert.match(pptManifest, /https:\/\/app\.decken\.ai\/addin\/icon-128\.png/);

  // 2. Validate WPS Office ribbon.xml
  const wpsRibbon = await readFile(new URL('../wps/ribbon.xml', import.meta.url), 'utf8');
  assert.match(wpsRibbon, /http:\/\/schemas\.kingsoft\.com\/office\/2009\/07\/customui/);
  assert.match(wpsRibbon, /OnToggleDeckenTaskPane/);

  // 3. Origin safety check
  assert.throws(() => validateApiOrigin('http://remote.attacker.example', true), /Refusing non-HTTPS/);
  assert.equal(validateApiOrigin('https://api.decken.ai', true), 'https://api.decken.ai');

  // 4. Mock generateDeckViaDecken (/v1/decks) & reconstructSlideViaDecken + Microsoft PowerPoint insertion
  const mockB64 = Buffer.from('PK\x03\x04office-slide').toString('base64');
  const deckJob = await generateDeckViaDecken({
    outline: '## Slide 1\n- Metric',
    apiKey: 'dk_office_test',
    fetchImpl: async (url, init) => {
      assert.equal(url, 'https://api.decken.ai/v1/decks');
      assert.equal(init.headers.Authorization, 'Bearer dk_office_test');
      return {
        ok: true,
        status: 200,
        json: async () => ({ jobId: 'deck_off_1', scenario: 'from_outline', status: 'completed', pptxBase64: mockB64 }),
      };
    },
  });
  assert.equal(deckJob.scenario, 'from_outline');

  const job = await reconstructSlideViaDecken({
    imageDataUrl: 'data:image/png;base64,iVBORw0KGgo=',
    apiKey: 'dk_office_test',
    fetchImpl: async (url, init) => {
      assert.equal(url, 'https://api.decken.ai/v1/reconstruct');
      assert.equal(init.headers.Authorization, 'Bearer dk_office_test');
      return {
        ok: true,
        status: 200,
        json: async () => ({ jobId: 'job_off_1', status: 'completed', pptxBase64: mockB64 }),
      };
    },
  });
  assert.equal(job.pptxBase64, mockB64);

  let insertedPptxCall = null;
  await insertIntoMicrosoftPowerPoint(mockB64, {
    powerPointGlobal: {
      InsertSlideFormatting: { keepSourceFormatting: 'KeepSourceFormatting' },
      run: async (cb) => {
        await cb({
          presentation: {
            getSelectedSlides: () => ({ items: [{ id: '267#' }], load: () => {} }),
            insertSlidesFromBase64: (b64, opts) => {
              insertedPptxCall = { b64, opts };
            },
          },
          sync: async () => {},
        });
      },
    },
  });
  assert.equal(insertedPptxCall.b64, mockB64);
  assert.equal(insertedPptxCall.opts.targetSlideId, '267#');

  // 5. Mock WPS Presentation JSAPI insertion
  let writtenFile = null;
  let wpsInserted = null;
  const wpsResult = insertIntoWpsPresentation(mockB64, {
    wpsGlobal: {
      Env: { GetTempPath: () => '/tmp/wps-test' },
      FileSystem: {
        writeAsBinaryString: (path, data) => {
          writtenFile = { path, len: data.length };
        },
      },
      WpsApplication: () => ({
        ActiveWindow: { Selection: { SlideRange: { Item: () => ({ SlideIndex: 3 }) } } },
        ActivePresentation: {
          Slides: {
            Count: 5,
            InsertFromFile: (file, idx) => {
              wpsInserted = { file, idx };
              return 1;
            },
          },
        },
      }),
    },
  });
  assert.ok(writtenFile.path.startsWith('/tmp/wps-test/decken_reconstruct_'));
  assert.equal(wpsInserted.idx, 3);
  assert.equal(wpsResult.insertedCount, 1);
});
