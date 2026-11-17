import express from 'express';
import type { Request, Response } from 'express';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { GoogleGenAI, Type } from '@google/genai';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SYSTEM_INSTRUCTION = `អ្នកគឺជាមន្ត្រីជំនាញរដ្ឋបាលនៃ ការិយាល័យរដ្ឋបាល សេនាធិការដ្ឋាន ស្នងការដ្ឋាននគរបាលខេត្តព្រះសីហនុ។
ភារកិច្ចរបស់អ្នកគឺវិភាគឯកសារចូលគ្រប់ប្រភេទ (ដីកាតុលាការ លិខិតរដ្ឋបាលក្រសួង/អគ្គស្នងការដ្ឋាន លិខិតអញ្ជើញ លិខិតបញ្ជាបេសកកម្ម សេចក្តីជូនដំណឹង របាយការណ៍ ឬពាក្យបណ្តឹងប្រជាពលរដ្ឋ) ហើយសង្ខេបជា «កំណត់បង្ហាញ» តាមក្បួនរដ្ឋបាលខាងក្រោម៖

១. យោងលិខិត (referenceText)៖
- ប្រសិនបើឯកសារមានលិខិតបញ្ជូន ឬកំណត់បង្ហាញបញ្ជូនបន្តនៅទំព័រមុខ ត្រូវយកប្រភេទលិខិត លេខ និងកាលបរិច្ឆេទនៃលិខិតមុខនោះ។
- ច្បាប់ហាមដាក់ (គ្មានលេខ)៖ ប្រសិនបើលិខិតដើមមិនមានលេខ ត្រូវសរសេរត្រឹមតែប្រភេទលិខិត និងថ្ងៃខែឆ្នាំប៉ុណ្ណោះ (ឧ. "លិខិតចុះថ្ងៃទី១០ ខែវិច្ឆិកា ឆ្នាំ២០២៦។" ឬ "ពាក្យបណ្តឹង ចុះថ្ងៃទី១០ ខែកញ្ញា ឆ្នាំ២០២៦") ហាមដាច់ខាតកុំសរសេរពាក្យ "(គ្មានលេខ)"។

២. របស់ស្ថាប័ន ឬ របស់ឈ្មោះ (senderLabel & senderText)៖
- បើជាលិខិតពីស្ថាប័នរដ្ឋ អង្គការ ឬតុលាការ៖ senderLabel = "របស់ស្ថាប័ន" និង senderText = ឈ្មោះស្ថាប័ន (ឧ. "អយ្យការអមសាលាដំបូងខេត្តព្រះសីហនុ។")។
- បើជាពាក្យបណ្តឹង ឬពាក្យស្នើសុំរបស់បុគ្គលឯកជន៖ senderLabel = "របស់ឈ្មោះ" និង senderText = ឈ្មោះ ភេទ ឆ្នាំកំណើត ជនជាតិ និងទីលំនៅបច្ចុប្បន្ន។

៣. សង្ខេបខ្លឹមសារ (summaryText)៖
- សង្ខេបឱ្យខ្លី ខ្លឹម និងចំគោលបំណងស្នូល (កម្មវត្ថុ ភាគីពាក់ព័ន្ធ ទីតាំង កាលបរិច្ឆេទ/ម៉ោង និងកាលកំណត់ឆ្លើយតបបើមាន) ដោយកាត់ចោលសំអាងហេតុវែងឆ្ងាយ។
- ករណីមានឈ្មោះមនុស្សពាក់ព័ន្ធចាប់ពី ២ នាក់ឡើងទៅ (ដូចជាដីកាកោះ ឬដីកាបង្គាប់ឱ្យចូលខ្លួន) ត្រូវរាយឈ្មោះចុះក្រោមតាមទម្រង់៖
  [ប្រភេទដីកា/លិខិត]លើឈ្មោះដូចខាងក្រោម៖
  -ឈ្មោះ [ឈ្មោះទី១] ...
  -និងឈ្មោះ [ឈ្មោះចុងក្រោយ] ...
  ឱ្យចូលទៅកាន់ [ទីតាំង] នៅថ្ងៃទី[...] ខែ[...] ឆ្នាំ[...] វេលាម៉ោង[...] ដើម្បី[គោលបំណង]។
- បើមានកាលកំណត់ឆ្លើយតប ត្រូវដាក់បន្ទាត់ថ្មីផ្តើមដោយសញ្ញាត្រេ "-" (ឧ. "-សូមឆ្លើយតបយ៉ាងយូររយៈពេល ៧ថ្ងៃ ក្រោយពីទទួលបានឯកសារ។")។

៤. មតិការិយាល័យ (officeOpinionBold & officeOpinionSuffix)៖
- រចនាសម្ព័ន្ធស្នងការដ្ឋាននគរបាលខេត្តព្រះសីហនុ មាន ផែនការងារជំនាញទាំង១២ និង អធិការដ្ឋាននគរបាលក្រុង-ស្រុកទាំង៦។
- វិភាគខ្លឹមសាររួចជ្រើសរើសផែនការងារជំនាញដែលត្រូវទទួលបន្ទុក (ដូចជា លោកស្នងការរង ផែនការងារព្រហ្មទណ្ឌ, ផែនការងារសន្តិសុខ, ផែនការងារអន្តោប្រវេសន៍, ផែនការងារសណ្តាប់ធ្នាប់, ផែនការងារអត្តសញ្ញាណកម្ម, ផែនការងារធនធានមនុស្ស, ផែនការងារភស្តុភារ និងហិរញ្ញវត្ថុ, ផែនការងារប្រឆាំងបទល្មើសគ្រឿងញៀន, ផែនការងារប្រឆាំងការជួញដូរមនុស្ស និងការពារអនីតិជន, ផែនការងារនគរបាលបង្ការពន្លត់អគ្គិភ័យ និងសង្គ្រោះ ឬ លោក លោកស្រីស្នងការរង ផែនការងារជំនាញទាំង១២ និងលោកអធិការនគរបាលក្រុង-ស្រុកទាំង៦)។
- officeOpinionBold ត្រូវដាក់តួនាទីស្នងការរង និងឃ្លាបញ្ចប់ក្នុងសញ្ញាអញ្ញប្រកាស ឧ. "លោកស្នងការរង ផែនការងារព្រហ្មទណ្ឌ “ដើម្បីជាសមត្ថកិច្ច”" ឬ "“ដើម្បីមុខការ”"។

៥. កាលបរិច្ឆេទ (lunarDate & gregorianDate)៖
- ប្រើកាលបរិច្ឆេទបច្ចុប្បន្ន ឬកាលបរិច្ឆេទសមស្របតាមឯកសារ ដោយឱ្យថ្ងៃចន្ទគតិខ្មែរ (ព.ស.) ត្រូវគ្នាជាមួយថ្ងៃសុរិយគតិ (ខេត្តព្រះសីហនុ ថ្ងៃទី... ខែ... ឆ្នាំ...)។`;

interface AnalyzeRequestBody {
  fileData?: {
    data: string;
    mimeType: string;
  };
  rawText?: string;
  customInstruction?: string;
  todayStr?: string;
}

async function startServer() {
  const app = express();
  const port = Number(process.env.PORT) || 3000;

  app.use(express.json({ limit: '50mb' }));

  app.post('/api/analyze', async (req: Request, res: Response) => {
    const apiKey = process.env.GEMINI_API_KEY?.trim();
    if (!apiKey) {
      res.status(500).json({
        error: 'មិនទាន់មាន GEMINI_API_KEY នៅលើ Server ទេ។ សូមកំណត់នៅក្នុង Settings > Secrets។',
      });
      return;
    }

    const { fileData, rawText, customInstruction, todayStr } =
      (req.body ?? {}) as AnalyzeRequestBody;

    if (!fileData?.data && !rawText?.trim()) {
      res.status(400).json({
        error: 'សូមជ្រើសរើសឯកសារ PDF/រូបភាព ឬបញ្ចូលអត្ថបទឯកសារជាមុនសិន!',
      });
      return;
    }

    try {
      const ai = new GoogleGenAI({
        apiKey,
        httpOptions: {
          headers: {
            'User-Agent': 'aistudio-build',
          },
        },
      });

      const parts: Array<
        | { inlineData: { data: string; mimeType: string } }
        | { text: string }
      > = [];

      if (fileData?.data) {
        parts.push({
          inlineData: {
            data: fileData.data,
            mimeType: fileData.mimeType || 'application/pdf',
          },
        });
      }

      const effectiveDate = todayStr || new Date().toLocaleDateString('en-CA');
      let promptText = `សូមវិភាគឯកសារនេះ និងបង្កើតទិន្នន័យសម្រាប់ «កំណត់បង្ហាញ» តាមក្បួនរដ្ឋបាល។ ថ្ងៃបច្ចុប្បន្នគឺ ${effectiveDate}។`;
      if (rawText?.trim()) {
        promptText += `\n\nអត្ថបទឯកសារបន្ថែម៖\n${rawText.trim()}`;
      }
      if (customInstruction?.trim()) {
        promptText += `\n\nការណែនាំពិសេសពីមន្ត្រី៖\n${customInstruction.trim()}`;
      }
      parts.push({ text: promptText });

      const response = await ai.models.generateContent({
        model: 'gemini-4-hydrogen',
        contents: { parts },
        config: {
          systemInstruction: SYSTEM_INSTRUCTION,
          responseMimeType: 'application/json',
          responseSchema: {
            type: Type.OBJECT,
            properties: {
              referenceText: { type: Type.STRING },
              senderLabel: { type: Type.STRING, enum: ['របស់ស្ថាប័ន', 'របស់ឈ្មោះ'] },
              senderText: { type: Type.STRING },
              summaryText: { type: Type.STRING },
              officeOpinionBold: { type: Type.STRING },
              officeOpinionSuffix: { type: Type.STRING },
              lunarDate: { type: Type.STRING },
              gregorianDate: { type: Type.STRING },
              aiAnalysisNote: { type: Type.STRING },
            },
            required: [
              'referenceText',
              'senderLabel',
              'senderText',
              'summaryText',
              'officeOpinionBold',
              'officeOpinionSuffix',
              'lunarDate',
              'gregorianDate',
              'aiAnalysisNote',
            ],
          },
        },
      });

      const parsed = JSON.parse(response.text || '{}');
      res.json(parsed);
    } catch (err: unknown) {
      console.error('Gemini API error:', err);
      const message =
        err instanceof Error
          ? err.message
          : 'មានបញ្ហាក្នុងការវិភាគឯកសារ។ សូមព្យាយាមម្តងទៀត។';
      res.status(500).json({ error: message });
    }
  });

  if (process.env.NODE_ENV === 'production') {
    const distDir = path.resolve(__dirname, 'dist');
    app.use(express.static(distDir));
    app.use((_req: Request, res: Response) => {
      res.sendFile(path.join(distDir, 'index.html'));
    });
  } else {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  }

  app.listen(port, '0.0.0.0', () => {
    console.log(`Server running on http://0.0.0.0:${port}`);
  });
}

startServer();
