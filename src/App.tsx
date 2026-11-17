import React, { useState, useRef } from 'react';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  TabStopType,
  BorderStyle,
  HeightRule,
} from 'docx';
import { saveAs } from 'file-saver';
import {
  FileUp,
  Sparkles,
  Download,
  Printer,
  FileText,
  AlertCircle,
  CheckCircle2,
  Loader2,
  Trash2,
} from 'lucide-react';

interface KumnatData {
  referenceText: string;
  senderLabel: 'របស់ស្ថាប័ន' | 'របស់ឈ្មោះ';
  senderText: string;
  summaryText: string;
  officeOpinionBold: string;
  officeOpinionSuffix: string;
  lunarDate: string;
  gregorianDate: string;
  aiAnalysisNote: string;
}

const VIP_REGEX = /(ឯកឧត្តម[^\s,។]+(?:\s+[^\s,។]+){1,7}|សម្តេច[^\s,។]+(?:\s+[^\s,។]+){1,7}|លោកជំទាវ[^\s,។]+(?:\s+[^\s,។]+){1,6})/g;

function buildKhmerRuns(text: string, defaultFont = 'Khmer OS Siemreap', defaultBold = false, sizeHalfPt = 24): TextRun[] {
  const parts = text.split(VIP_REGEX);
  const runs: TextRun[] = [];
  for (const part of parts) {
    if (!part) continue;
    const isVip =
      part.startsWith('ឯកឧត្តម') ||
      part.startsWith('សម្តេច') ||
      part.startsWith('លោកជំទាវ');
    const fontName = isVip ? 'Khmer M1' : defaultFont;
    runs.push(
      new TextRun({
        text: part,
        bold: isVip ? false : defaultBold,
        size: sizeHalfPt,
        font: {
          name: fontName,
          ascii: fontName,
          hAnsi: fontName,
          cs: fontName,
          eastAsia: fontName,
        },
      })
    );
  }
  return runs;
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [rawText, setRawText] = useState<string>('');
  const [customInstruction, setCustomInstruction] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const [data, setData] = useState<KumnatData>({
    referenceText: 'ដីកាបញ្ជូនលេខ ៧៨៩៥ អ.យ.ដ/២៦ ចុះថ្ងៃទី២៨ ខែតុលា ឆ្នាំ២០២៦។',
    senderLabel: 'របស់ស្ថាប័ន',
    senderText: 'អយ្យការអមសាលាដំបូងខេត្តព្រះសីហនុ។',
    summaryText:
      'សំណើសុំចាត់អញ្ជើញតំណាងស្នងការ នៃស្នងការដ្ឋាននគរបាលខេត្ត ចូលរួមក្នុងពិធីប្រកាសសមាសភាពគណៈកម្មាធិការប្រយុទ្ធនឹងជំងឺអេដស៍ និងកាមរោគ ក្រសួងមហាផ្ទៃ ក្រោមអធិបតីភាពដ៏ខ្ពង់ខ្ពស់ ឯកឧត្តមអភិសន្តិបណ្ឌិត ស សុខា ឧបនាយករដ្ឋមន្ត្រី រដ្ឋមន្ត្រីក្រសួងមហាផ្ទៃ។',
    officeOpinionBold: 'លោកស្នងការរង ផែនការងារធនធានមនុស្ស “ដើម្បីមុខការ”',
    officeOpinionSuffix: '។',
    lunarDate: 'ថ្ងៃពុធ ៩កើត ខែកត្តិក ឆ្នាំមមី អដ្ឋស័ក ព.ស.២៥៧០',
    gregorianDate: 'ខេត្តព្រះសីហនុ ថ្ងៃទី១៨ ខែវិច្ឆិកា ឆ្នាំ២០២៦',
    aiAnalysisNote: 'រួចរាល់សម្រាប់បញ្ចូលឯកសារថ្មីដើម្បីវិភាគដោយ AI។',
  });

  const fileToBase64 = (f: File): Promise<string> =>
    new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const res = reader.result as string;
        resolve(res.split(',')[1]);
      };
      reader.onerror = reject;
      reader.readAsDataURL(f);
    });

  const handleAnalyze = async () => {
    if (!file && !rawText.trim()) {
      setError('សូមជ្រើសរើសឯកសារ PDF/រូបភាព ឬបញ្ចូលអត្ថបទឯកសារជាមុនសិន!');
      return;
    }

    setLoading(true);
    setError('');

    try {
      let fileData: { data: string; mimeType: string } | undefined;
      if (file) {
        const b64 = await fileToBase64(file);
        fileData = {
          data: b64,
          mimeType: file.type || 'application/pdf',
        };
      }

      const todayStr = new Date().toLocaleDateString('en-CA');
      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileData,
          rawText: rawText.trim(),
          customInstruction: customInstruction.trim(),
          todayStr,
        }),
      });

      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload?.error || 'មានបញ្ហាក្នុងការវិភាគឯកសារ។ សូមព្យាយាមម្តងទៀត។');
      }

      setData(payload as KumnatData);
    } catch (err: unknown) {
      console.error(err);
      const message =
        err instanceof Error
          ? err.message
          : 'មានបញ្ហាក្នុងការវិភាគឯកសារ។ សូមពិនិត្យឯកសារម្តងទៀត។';
      setError(message);
    } finally {
      setLoading(false);
    }
  };

  const handleDownloadDocx = async () => {
    const tabStops = [
      { type: TabStopType.LEFT, position: 2160 }, // 1.5 in
      { type: TabStopType.LEFT, position: 2448 }, // 1.7 in
    ];

    const makeRowParagraph = (label: string, valueRuns: TextRun[]) =>
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        tabStops,
        spacing: { before: 40, after: 40, line: 276 },
        children: [
          new TextRun({
            text: `${label}	:	`,
            bold: label === 'សង្ខេបខ្លឹមសារ',
            size: 24,
            font: {
              name: 'Khmer OS Siemreap',
              ascii: 'Khmer OS Siemreap',
              hAnsi: 'Khmer OS Siemreap',
              cs: 'Khmer OS Siemreap',
              eastAsia: 'Khmer OS Siemreap',
            },
          }),
          ...valueRuns,
        ],
      });

    // Build summary lines properly with tab indents on multi-line lists
    const summaryLines = data.summaryText.split('\n');
    const summaryRuns: TextRun[] = [];
    summaryLines.forEach((line, idx) => {
      const cleanLine = line.trim();
      if (!cleanLine) return;
      if (idx > 0) {
        summaryRuns.push(
          new TextRun({
            break: 1,
            text: '		',
            size: 24,
            font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
          })
        );
      }
      summaryRuns.push(...buildKhmerRuns(cleanLine, 'Khmer OS Siemreap', false, 24));
    });

    const headerTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE },
        bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE },
        right: { style: BorderStyle.NONE },
        insideHorizontal: { style: BorderStyle.NONE },
        insideVertical: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 52, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20 },
                  children: [
                    new TextRun({
                      text: 'ស្នងការដ្ឋាននគរបាលខេត្តព្រះសីហនុ',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20 },
                  children: [
                    new TextRun({
                      text: 'សេនាធិការដ្ឋាន',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20 },
                  children: [
                    new TextRun({
                      text: 'ការិយាល័យរដ្ឋបាល',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              width: { size: 48, type: WidthType.PERCENTAGE },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20 },
                  children: [
                    new TextRun({
                      text: 'ព្រះរាជាណាចក្រកម្ពុជា',
                      size: 23,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20 },
                  children: [
                    new TextRun({
                      text: 'ជាតិ សាសនា ព្រះមហាក្សត្រ',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    });

    const bottomTable = new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          height: { value: 3600, rule: HeightRule.ATLEAST }, // ~180pt tall box for generous handwriting
          children: [
            new TableCell({
              width: { size: 47, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.NONE },
                bottom: { style: BorderStyle.NONE },
                left: { style: BorderStyle.NONE },
                right: { style: BorderStyle.SINGLE, size: 6, color: '000000' },
              },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { before: 80, after: 20 },
                  children: [
                    new TextRun({
                      text: 'អគ្គស្នងការរង និងជា',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [
                    new TextRun({
                      text: 'ស្នងការនគរបាលខេត្តព្រះសីហនុ',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
              ],
            }),
            new TableCell({
              width: { size: 53, type: WidthType.PERCENTAGE },
              borders: {
                top: { style: BorderStyle.SINGLE, size: 6, color: '000000' },
                bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000' },
                left: { style: BorderStyle.SINGLE, size: 6, color: '000000' },
                right: { style: BorderStyle.SINGLE, size: 6, color: '000000' },
              },
              margins: { top: 100, bottom: 100, left: 140, right: 140 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20 },
                  children: [
                    new TextRun({
                      text: 'ស្នងការរង',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 60 },
                  children: [
                    new TextRun({
                      text: 'ផែនការងារសេនាធិការ',
                      size: 22,
                      font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                    }),
                  ],
                }),
                new Paragraph({
                  spacing: { after: 60 },
                  children: [
                    new TextRun({
                      text: 'គោរពជូន : ឯកឧត្តម ឧត្តមសេនីយ៍ឯក អគ្គស្នងការរង\n               និងជាស្នងការនគរបាលខេត្តព្រះសីហនុ',
                      size: 21,
                      font: {
                        name: 'Khmer OS Siemreap',
                        ascii: 'Khmer OS Siemreap',
                        hAnsi: 'Khmer OS Siemreap',
                        cs: 'Khmer OS Siemreap',
                      },
                    }),
                  ],
                }),
                new Paragraph({
                  children: [
                    new TextRun({
                      text: 'គោរពជូនយោបល់ :',
                      size: 21,
                      font: {
                        name: 'Khmer OS Siemreap',
                        ascii: 'Khmer OS Siemreap',
                        hAnsi: 'Khmer OS Siemreap',
                        cs: 'Khmer OS Siemreap',
                      },
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    });

    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: { width: 11906, height: 16838 }, // A4
              margin: { top: 700, bottom: 260, left: 980, right: 650 },
            },
          },
          children: [
            headerTable,
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 120, after: 140 },
              children: [
                new TextRun({
                  text: 'កំណត់បង្ហាញ',
                  size: 26,
                  font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                }),
              ],
            }),
            makeRowParagraph('យោងលិខិត', buildKhmerRuns(data.referenceText)),
            makeRowParagraph(data.senderLabel, buildKhmerRuns(data.senderText)),
            makeRowParagraph('សង្ខេបខ្លឹមសារ', summaryRuns),
            makeRowParagraph('មតិការិយាល័យ ', [
              new TextRun({
                text: 'សូមគោរពអនុញ្ញាតគួរប្រគល់ជូន ',
                size: 24,
                font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
              }),
              new TextRun({
                text: data.officeOpinionBold,
                bold: true,
                size: 24,
                font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
              }),
              new TextRun({
                text: data.officeOpinionSuffix,
                size: 24,
                font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
              }),
            ]),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              tabStops,
              spacing: { before: 80, after: 40, line: 276 },
              children: [
                new TextRun({
                  text: '		សូមគោរពជូន ',
                  size: 24,
                  font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
                }),
                new TextRun({
                  text: 'ឯកឧត្តម ឧត្តមសេនីយ៍ឯក អគ្គស្នងការរង និងជាស្នងការនគរបាលខេត្តព្រះសីហនុ',
                  size: 24,
                  font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                }),
                new TextRun({
                  text: ' មេត្តាជ្រាបនិងសម្រេចដ៏ខ្ពង់ខ្ពស់។',
                  size: 24,
                  font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 60 },
              children: [
                new TextRun({
                  text: 'ដោយក្តីគោរព។',
                  size: 24,
                  font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { after: 20 },
              children: [
                new TextRun({
                  text: data.lunarDate,
                  size: 24,
                  font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { after: 40 },
              children: [
                new TextRun({
                  text: data.gregorianDate,
                  size: 24,
                  font: { name: 'Khmer OS Siemreap', cs: 'Khmer OS Siemreap' },
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { after: 720 }, // Generous space for នាយការិយាល័យ signature & stamp
              children: [
                new TextRun({
                  text: 'នាយការិយាល័យ		',
                  size: 24,
                  font: { name: 'Khmer M1', ascii: 'Khmer M1', hAnsi: 'Khmer M1', cs: 'Khmer M1' },
                }),
              ],
            }),
            bottomTable,
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `កំណត់បង្ហាញ_${new Date().toISOString().slice(0, 10)}.docx`);
  };

  return (
    <div className="min-h-screen bg-slate-100">
      {/* Top Navbar */}
      <header className="no-print bg-blue-950 text-white shadow-md">
        <div className="max-w-7xl mx-auto px-4 py-3 flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <div className="bg-amber-500/20 border border-amber-400/40 p-2 rounded-lg">
              <FileText className="w-6 h-6 text-amber-300" />
            </div>
            <div>
              <h1 className="font-khmer-m1 text-base md:text-lg text-amber-300">
                ប្រព័ន្ធបង្កើត «កំណត់បង្ហាញ» ស្វ័យប្រវត្តិ (AI)
              </h1>
              <p className="text-xs text-blue-200">
                ការិយាល័យរដ្ឋបាល • សេនាធិការដ្ឋាន • ស្នងការដ្ឋាននគរបាលខេត្តព្រះសីហនុ
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => window.print()}
              className="flex items-center gap-1.5 text-xs bg-slate-700 hover:bg-slate-600 px-3 py-2 rounded-lg transition"
            >
              <Printer className="w-4 h-4" />
              <span>បោះពុម្ព</span>
            </button>
            <button
              onClick={handleDownloadDocx}
              className="flex items-center gap-1.5 text-xs bg-emerald-600 hover:bg-emerald-500 font-semibold px-4 py-2 rounded-lg shadow transition"
            >
              <Download className="w-4 h-4" />
              <span>ទាញយក Word (.docx)</span>
            </button>
          </div>
        </div>
      </header>

      {/* Main Split Workspace */}
      <main className="max-w-7xl mx-auto p-4 grid grid-cols-1 lg:grid-cols-12 gap-6">
        {/* Left Control & AI Analyzer Panel */}
        <section className="no-print lg:col-span-5 space-y-4">
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 space-y-4">
            <h2 className="font-bold text-sm text-blue-950 flex items-center gap-2 border-b pb-2">
              <Sparkles className="w-4 h-4 text-amber-500" />
              <span>១. បញ្ចូលឯកសារដើម្បីឱ្យ AI វិភាគ</span>
            </h2>

            <div
              onClick={() => fileInputRef.current?.click()}
              className="border-2 border-dashed border-blue-300 hover:border-blue-500 bg-blue-50/50 rounded-xl p-5 text-center cursor-pointer transition"
            >
              <input
                ref={fileInputRef}
                type="file"
                accept=".pdf,image/*"
                className="hidden"
                onChange={(e) => setFile(e.target.files?.[0] || null)}
              />
              <FileUp className="w-8 h-8 text-blue-600 mx-auto mb-2" />
              {file ? (
                <div className="flex items-center justify-center gap-2 text-xs font-semibold text-blue-900">
                  <span>{file.name}</span>
                  <button
                    onClick={(e) => {
                      e.stopPropagation();
                      setFile(null);
                    }}
                    className="text-red-500 hover:text-red-700"
                  >
                    <Trash2 className="w-4 h-4" />
                  </button>
                </div>
              ) : (
                <>
                  <p className="text-xs font-semibold text-slate-700">
                    ចុចទីនេះដើម្បីជ្រើសរើសឯកសារ PDF ឬ រូបភាពស្កេន
                  </p>
                  <p className="text-[11px] text-slate-500 mt-1">
                    គាំទ្រដីកាតុលាការ លិខិតរដ្ឋបាល ពាក្យបណ្តឹង របាយការណ៍គ្រប់ប្រភេទ
                  </p>
                </>
              )}
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                ឬ បិទភ្ជាប់អត្ថបទឯកសារ (ស្រេចចិត្ត)៖
              </label>
              <textarea
                rows={3}
                value={rawText}
                onChange={(e) => setRawText(e.target.value)}
                placeholder="ចម្លងខ្លឹមសារលិខិតដាក់ទីនេះ ប្រសិនបើមិនមានឯកសារ PDF/រូបភាព..."
                className="w-full text-xs border border-slate-300 rounded-lg p-2.5 focus:ring-2 focus:ring-blue-500 outline-none"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-700 mb-1">
                ការណែនាំបន្ថែមសម្រាប់ AI (ស្រេចចិត្ត)៖
              </label>
              <input
                type="text"
                value={customInstruction}
                onChange={(e) => setCustomInstruction(e.target.value)}
                placeholder="ឧ. ប្រគល់ជូនផែនការងារអន្តោប្រវេសន៍ ឬ បញ្ជាក់កាលកំណត់ ៥ថ្ងៃ..."
                className="w-full text-xs border border-slate-300 rounded-lg px-3 py-2 focus:ring-2 focus:ring-blue-500 outline-none"
              />
            </div>

            {error && (
              <div className="bg-red-50 border border-red-200 text-red-700 text-xs p-3 rounded-lg flex items-start gap-2">
                <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                <span>{error}</span>
              </div>
            )}

            <button
              onClick={handleAnalyze}
              disabled={loading}
              className="w-full bg-blue-900 hover:bg-blue-800 disabled:bg-blue-400 text-white font-semibold text-sm py-2.5 rounded-xl shadow flex items-center justify-center gap-2 transition"
            >
              {loading ? (
                <>
                  <Loader2 className="w-4 h-4 animate-spin" />
                  <span>AI កំពុងវិភាគ និងសង្ខេបខ្លឹមសារ...</span>
                </>
              ) : (
                <>
                  <Sparkles className="w-4 h-4 text-amber-300" />
                  <span>វិភាគ និងបង្កើតកំណត់បង្ហាញដោយ AI</span>
                </>
              )}
            </button>

            {data.aiAnalysisNote && (
              <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs p-2.5 rounded-lg flex items-start gap-2">
                <CheckCircle2 className="w-4 h-4 shrink-0 mt-0.5 text-emerald-600" />
                <span>{data.aiAnalysisNote}</span>
              </div>
            )}
          </div>

          {/* Manual Quick Edit Fields */}
          <div className="bg-white rounded-xl shadow-sm border border-slate-200 p-4 space-y-3">
            <h2 className="font-bold text-sm text-blue-950 border-b pb-2">
              ២. ពិនិត្យ និងកែសម្រួលខ្លឹមសារមុនទាញយក Word
            </h2>

            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">យោងលិខិត៖</label>
              <input
                type="text"
                value={data.referenceText}
                onChange={(e) => setData({ ...data, referenceText: e.target.value })}
                className="w-full text-xs border border-slate-300 rounded-lg px-3 py-1.5"
              />
            </div>

            <div className="grid grid-cols-3 gap-2">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">ប្រភេទអ្នកផ្ញើ៖</label>
                <select
                  value={data.senderLabel}
                  onChange={(e) =>
                    setData({ ...data, senderLabel: e.target.value as 'របស់ស្ថាប័ន' | 'របស់ឈ្មោះ' })
                  }
                  className="w-full text-xs border border-slate-300 rounded-lg px-2 py-1.5"
                >
                  <option value="របស់ស្ថាប័ន">របស់ស្ថាប័ន</option>
                  <option value="របស់ឈ្មោះ">របស់ឈ្មោះ</option>
                </select>
              </div>
              <div className="col-span-2">
                <label className="block text-xs font-semibold text-slate-600 mb-1">
                  ឈ្មោះស្ថាប័ន ឬ ឈ្មោះដើមបណ្តឹង៖
                </label>
                <input
                  type="text"
                  value={data.senderText}
                  onChange={(e) => setData({ ...data, senderText: e.target.value })}
                  className="w-full text-xs border border-slate-300 rounded-lg px-3 py-1.5"
                />
              </div>
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">សង្ខេបខ្លឹមសារ៖</label>
              <textarea
                rows={4}
                value={data.summaryText}
                onChange={(e) => setData({ ...data, summaryText: e.target.value })}
                className="w-full text-xs border border-slate-300 rounded-lg p-2.5 leading-relaxed"
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 mb-1">
                មតិការិយាល័យ (ផែនការងារទទួលបន្ទុក)៖
              </label>
              <input
                type="text"
                value={data.officeOpinionBold}
                onChange={(e) => setData({ ...data, officeOpinionBold: e.target.value })}
                className="w-full text-xs border border-slate-300 rounded-lg px-3 py-1.5 font-semibold"
              />
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">ថ្ងៃចន្ទគតិ៖</label>
                <input
                  type="text"
                  value={data.lunarDate}
                  onChange={(e) => setData({ ...data, lunarDate: e.target.value })}
                  className="w-full text-xs border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
              <div>
                <label className="block text-xs font-semibold text-slate-600 mb-1">ថ្ងៃសុរិយគតិ៖</label>
                <input
                  type="text"
                  value={data.gregorianDate}
                  onChange={(e) => setData({ ...data, gregorianDate: e.target.value })}
                  className="w-full text-xs border border-slate-300 rounded-lg px-2.5 py-1.5"
                />
              </div>
            </div>
          </div>
        </section>

        {/* Right Live A4 Preview */}
        <section className="lg:col-span-7 flex justify-center overflow-x-auto">
          <div
            className="a4-sheet bg-white shadow-xl border border-slate-300 text-black flex flex-col justify-between"
            style={{
              width: '210mm',
              minHeight: '297mm',
              padding: '13mm 14mm 10mm 18mm',
              fontSize: '12pt',
              lineHeight: 1.55,
            }}
          >
            <div>
              {/* Header */}
              <div className="flex justify-between items-start mb-2">
                <div className="text-center font-khmer-m1 text-[11pt] leading-snug">
                  <div>ស្នងការដ្ឋាននគរបាលខេត្តព្រះសីហនុ</div>
                  <div>សេនាធិការដ្ឋាន</div>
                  <div>ការិយាល័យរដ្ឋបាល</div>
                </div>
                <div className="text-center font-khmer-m1 text-[11.5pt] leading-snug">
                  <div>ព្រះរាជាណាចក្រកម្ពុជា</div>
                  <div>ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
                </div>
              </div>

              {/* Title */}
              <h2 className="text-center font-khmer-m1 text-[13pt] my-3">កំណត់បង្ហាញ</h2>

              {/* Main 4 Rows */}
              <div className="space-y-1.5 text-[12pt]">
                <div className="grid grid-cols-[135px_18px_1fr] items-baseline">
                  <span>យោងលិខិត</span>
                  <span>:</span>
                  <span className="text-justify">{data.referenceText}</span>
                </div>

                <div className="grid grid-cols-[135px_18px_1fr] items-baseline">
                  <span>{data.senderLabel}</span>
                  <span>:</span>
                  <span className="text-justify">{data.senderText}</span>
                </div>

                <div className="grid grid-cols-[135px_18px_1fr] items-baseline">
                  <span className="font-bold">សង្ខេបខ្លឹមសារ</span>
                  <span className="font-bold">:</span>
                  <div className="text-justify whitespace-pre-line">{data.summaryText}</div>
                </div>

                <div className="grid grid-cols-[135px_18px_1fr] items-baseline">
                  <span>មតិការិយាល័យ</span>
                  <span>:</span>
                  <div className="text-justify">
                    សូមគោរពអនុញ្ញាតគួរប្រគល់ជូន{' '}
                    <span className="font-bold">{data.officeOpinionBold}</span>
                    {data.officeOpinionSuffix}
                  </div>
                </div>
              </div>

              {/* Salutation */}
              <div className="mt-3 text-justify indent-[153px]">
                សូមគោរពជូន{' '}
                <span className="font-khmer-m1">
                  ឯកឧត្តម ឧត្តមសេនីយ៍ឯក អគ្គស្នងការរង និងជាស្នងការនគរបាលខេត្តព្រះសីហនុ
                </span>{' '}
                មេត្តាជ្រាបនិងសម្រេចដ៏ខ្ពង់ខ្ពស់។
              </div>
              <div className="text-center my-1">ដោយក្តីគោរព។</div>

              {/* Dates & Office Chief Signature Space */}
              <div className="flex justify-end mt-1">
                <div className="text-center min-w-[290px]">
                  <div>{data.lunarDate}</div>
                  <div>{data.gregorianDate}</div>
                  <div className="font-khmer-m1 mt-1 mb-16">នាយការិយាល័យ</div>
                </div>
              </div>
            </div>

            {/* Bottom Leadership Signature Table */}
            <div className="grid grid-cols-12 mt-2" style={{ minHeight: '190px' }}>
              <div className="col-span-5 text-center pt-3 pr-2">
                <div className="font-khmer-m1 text-[11pt] leading-snug">
                  <div>អគ្គស្នងការរង និងជា</div>
                  <div>ស្នងការនគរបាលខេត្តព្រះសីហនុ</div>
                </div>
              </div>
              <div className="col-span-7 border border-black p-3 flex flex-col justify-between">
                <div>
                  <div className="text-center font-khmer-m1 text-[11pt] leading-snug mb-2">
                    <div>ស្នងការរង</div>
                    <div>ផែនការងារសេនាធិការ</div>
                  </div>
                  <div className="text-[10.5pt] leading-snug">
                    <div>គោរពជូន : ឯកឧត្តម ឧត្តមសេនីយ៍ឯក អគ្គស្នងការរង</div>
                    <div className="pl-[72px]">និងជាស្នងការនគរបាលខេត្តព្រះសីហនុ</div>
                  </div>
                  <div className="text-[10.5pt] mt-2">គោរពជូនយោបល់ :</div>
                </div>
              </div>
            </div>
          </div>
        </section>
      </main>
    </div>
  );
}
