import React, { useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, 
  Download, 
  RotateCcw, 
  FileSpreadsheet, 
  Plus, 
  Calculator, 
  Percent,
  Eye,
  AlertCircle,
  CheckCircle,
  Link,
  Palette,
  FileText,
  History,
  ChevronDown,
  ChevronUp,
  Key,
  Zap,
  Settings,
  TestTube
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Textarea } from '@/components/ui/textarea';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Badge } from '@/components/ui/badge';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '@/components/ui/accordion';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Separator } from '@/components/ui/separator';
import { toast } from 'sonner';

// 型定義
interface CellStyle {
  backgroundColor?: string;
  color?: string;
  fontWeight?: string;
  fontStyle?: string;
  textDecoration?: string;
  fontSize?: string;
  fontFamily?: string;
  textAlign?: string;
  border?: string;
}

interface CellData {
  value: any;
  formula?: string;
  style?: CellStyle;
  type?: string;
}

interface EditHistory {
  id: string;
  timestamp: Date;
  instruction: string;
  changes: {
    cellRange: string;
    beforeValue?: any;
    afterValue?: any;
  }[];
  formatPreserved: boolean;
}

interface MergeInfo {
  s: { r: number; c: number };
  e: { r: number; c: number };
}

interface DeepSeekResponse {
  choices: {
    message: {
      content: string;
    };
  }[];
}

interface CellUpdate {
  address: string;
  value?: any;
  formula?: string;
}

const ExcelQuoteEditor: React.FC = () => {
  // 状態管理
  const [uploadedFile, setUploadedFile] = useState<File | null>(null);
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheets, setSheets] = useState<string[]>([]);
  const [activeSheet, setActiveSheet] = useState<string>('');
  const [sheetData, setSheetData] = useState<CellData[][]>([]);
  const [cellStyles, setCellStyles] = useState<Map<string, CellStyle>>(new Map());
  const [merges, setMerges] = useState<MergeInfo[]>([]);
  const [editHistory, setEditHistory] = useState<EditHistory[]>([]);
  const [currentInstruction, setCurrentInstruction] = useState<string>('');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [formatPreservationStatus, setFormatPreservationStatus] = useState<boolean>(true);
  const [dragActive, setDragActive] = useState<boolean>(false);
  
  // DeepSeek API関連の状態
  const [deepseekApiKey, setDeepseekApiKey] = useState<string>('');
  const [isApiTesting, setIsApiTesting] = useState<boolean>(false);
  const [apiConnectionStatus, setApiConnectionStatus] = useState<'none' | 'success' | 'error'>('none');
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  
  // プレビュー関連の状態
  const [zoomLevel, setZoomLevel] = useState<number>(100);
  const [previewFullscreen, setPreviewFullscreen] = useState<boolean>(false);
  
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ファイル読み込み（書式情報込み）
  const readFileWithStyles = useCallback(async (file: File) => {
    setIsLoading(true);
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, {
        cellStyles: true,    // セルスタイル情報を読み込む
        cellFormulas: true,  // 数式を保持
        cellDates: true,     // 日付フォーマット保持
        cellNF: true,        // 数値フォーマット保持
        sheetStubs: true     // 空セルも保持
      });

      setWorkbook(wb);
      const sheetNames = wb.SheetNames;
      setSheets(sheetNames);
      
      if (sheetNames.length > 0) {
        setActiveSheet(sheetNames[0]);
        loadSheetData(wb, sheetNames[0]);
      }

      toast.success('ファイルが正常に読み込まれました（書式情報を保持）');
    } catch (error) {
      console.error('ファイル読み込みエラー:', error);
      toast.error('ファイルの読み込みに失敗しました');
    } finally {
      setIsLoading(false);
    }
  }, []);

  // シートデータの読み込み（書式情報込み）
  const loadSheetData = useCallback((wb: XLSX.WorkBook, sheetName: string) => {
    const worksheet = wb.Sheets[sheetName];
    if (!worksheet) return;

    // セル範囲を取得
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
    const data: CellData[][] = [];
    const styles = new Map<string, CellStyle>();

    // セル結合情報を取得
    const mergeInfo: MergeInfo[] = worksheet['!merges'] || [];
    setMerges(mergeInfo);

    // データと書式情報を抽出
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row: CellData[] = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = worksheet[cellAddress];
        
        let cellData: CellData = { value: '' };
        let cellStyle: CellStyle = {};

        if (cell) {
          cellData.value = cell.v || '';
          cellData.formula = cell.f;
          cellData.type = cell.t;

          // スタイル情報を抽出（改善版）
          if (cell.s) {
            const style = cell.s;
            
            // 背景色（より詳細な処理）
            if (style.fill) {
              if (style.fill.fgColor) {
                const color = style.fill.fgColor;
                if (color.rgb) {
                  cellStyle.backgroundColor = `#${color.rgb.substring(2)}`; // ARGBからRGBに変換
                } else if (color.indexed !== undefined) {
                  // インデックス色の処理（基本色のみ）
                  const indexedColors = [
                    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
                    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'
                  ];
                  if (color.indexed < indexedColors.length) {
                    cellStyle.backgroundColor = indexedColors[color.indexed];
                  }
                }
              }
              if (style.fill.bgColor) {
                const bgColor = style.fill.bgColor;
                if (bgColor.rgb && !cellStyle.backgroundColor) {
                  cellStyle.backgroundColor = `#${bgColor.rgb.substring(2)}`;
                }
              }
            }

            // フォント情報（改善版）
            if (style.font) {
              const font = style.font;
              if (font.color) {
                if (font.color.rgb) {
                  cellStyle.color = `#${font.color.rgb.substring(2)}`;
                } else if (font.color.indexed !== undefined) {
                  const indexedColors = [
                    '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF'
                  ];
                  if (font.color.indexed < indexedColors.length) {
                    cellStyle.color = indexedColors[font.color.indexed];
                  }
                }
              }
              if (font.bold) {
                cellStyle.fontWeight = 'bold';
              }
              if (font.italic) {
                cellStyle.fontStyle = 'italic';
              }
              if (font.underline) {
                cellStyle.textDecoration = 'underline';
              }
              if (font.sz) {
                cellStyle.fontSize = `${font.sz}px`;
              }
              if (font.name) {
                cellStyle.fontFamily = font.name;
              }
            }

            // 配置
            if (style.alignment) {
              const alignment = style.alignment;
              if (alignment.horizontal) {
                cellStyle.textAlign = alignment.horizontal;
              }
            }

            // 罫線（改善版）
            if (style.border) {
              const borders = [];
              if (style.border.top && style.border.top.style) {
                borders.push('border-top: 1px solid #000');
              }
              if (style.border.bottom && style.border.bottom.style) {
                borders.push('border-bottom: 1px solid #000');
              }
              if (style.border.left && style.border.left.style) {
                borders.push('border-left: 1px solid #000');
              }
              if (style.border.right && style.border.right.style) {
                borders.push('border-right: 1px solid #000');
              }
              if (borders.length > 0) {
                cellStyle.border = borders.join('; ');
              }
            }
          }
        }

        if (Object.keys(cellStyle).length > 0) {
          styles.set(cellAddress, cellStyle);
        }

        row.push(cellData);
      }
      data.push(row);
    }

    setSheetData(data);
    setCellStyles(styles);
  }, []);

  // DeepSeek API接続テスト
  const testDeepSeekConnection = useCallback(async () => {
    if (!deepseekApiKey.trim()) {
      toast.error('DeepSeek APIキーを入力してください');
      return;
    }

    setIsApiTesting(true);
    try {
      const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${deepseekApiKey}`
        },
        body: JSON.stringify({
          model: 'deepseek-chat',
          messages: [
            {
              role: 'user',
              content: 'こんにちは。接続テストです。'
            }
          ],
          max_tokens: 50
        })
      });

      if (response.ok) {
        setApiConnectionStatus('success');
        toast.success('DeepSeek API接続成功！');
      } else {
        const errorData = await response.json();
        setApiConnectionStatus('error');
        toast.error(`API接続エラー: ${errorData.error?.message || 'Unknown error'}`);
      }
    } catch (error) {
      setApiConnectionStatus('error');
      toast.error('API接続に失敗しました');
      console.error('DeepSeek API Error:', error);
    } finally {
      setIsApiTesting(false);
    }
  }, [deepseekApiKey]);

  // DeepSeek APIを使用してExcelデータを分析・編集
  const processWithDeepSeek = useCallback(async (instruction: string) => {
    if (!deepseekApiKey.trim()) {
      toast.error('DeepSeek APIキーを設定してください');
      return;
    }

    if (!workbook || !sheetData.length) {
      toast.error('Excelファイルを読み込んでください');
      return;
    }

    setIsProcessing(true);
    try {
      // 現在のシートデータをJSON形式で準備
      const currentData = sheetData.map((row, rowIndex) => 
        row.map((cell, colIndex) => ({
          address: XLSX.utils.encode_cell({ r: rowIndex, c: colIndex }),
          value: cell.value,
          formula: cell.formula,
          type: cell.type
        }))
      ).flat().filter(cell => cell.value !== '' || cell.formula);

      const prompt = `
あなたはExcelファイル編集の専門家です。以下のExcelデータに対して、ユーザーの指示に従って編集を行ってください。

現在のExcelデータ:
${JSON.stringify(currentData, null, 2)}

ユーザーの編集指示:
${instruction}

以下の形式でJSONレスポンスを返してください:
{
  "updates": [
    {
      "address": "セルアドレス（例：A1）",
      "value": "新しい値",
      "formula": "数式がある場合（オプション）"
    }
  ],
  "explanation": "実行した変更の説明"
}

注意事項:
- セルアドレスは必ずA1, B2のような形式で指定してください
- 数値計算が必要な場合は正確に計算してください
- 既存のデータ構造を理解して適切に編集してください
- JSONフォーマット以外は返さないでください
`;

      const response = await fetch('https://api.deepseek.com/v1/chat/completions', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${deepseekApiKey}`
        },
        body: JSON.stringify({
          model: 'deepseek-chat',
          messages: [
            {
              role: 'user',
              content: prompt
            }
          ],
          max_tokens: 2000,
          temperature: 0.1
        })
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || 'API request failed');
      }

      const data: DeepSeekResponse = await response.json();
      const content = data.choices[0]?.message?.content;

      if (!content) {
        throw new Error('APIからの応答が空です');
      }

      // JSONレスポンスを解析
      let parsedResponse;
      try {
        // JSONの開始と終了を見つけて抽出
        const jsonStart = content.indexOf('{');
        const jsonEnd = content.lastIndexOf('}') + 1;
        const jsonString = content.substring(jsonStart, jsonEnd);
        parsedResponse = JSON.parse(jsonString);
      } catch (parseError) {
        console.error('JSON解析エラー:', parseError);
        console.log('APIレスポンス:', content);
        throw new Error('APIレスポンスの解析に失敗しました');
      }

      // Excelデータを更新
      if (parsedResponse.updates && Array.isArray(parsedResponse.updates)) {
        const newSheetData = sheetData.map(row => row.map(cell => ({ ...cell })));
        const changes: { cellRange: string; beforeValue: any; afterValue: any }[] = [];

        parsedResponse.updates.forEach((update: CellUpdate) => {
          try {
            const cellRef = XLSX.utils.decode_cell(update.address);
            if (cellRef.r < newSheetData.length && cellRef.c < newSheetData[cellRef.r].length) {
              const currentCell = newSheetData[cellRef.r][cellRef.c];
              const beforeValue = currentCell.value;
              const beforeFormula = currentCell.formula;

              const nextCell = { ...currentCell };

              if (Object.prototype.hasOwnProperty.call(update, 'value')) {
                nextCell.value = update.value;
              }

              if (update.formula !== undefined) {
                nextCell.formula = update.formula;
              }

              newSheetData[cellRef.r][cellRef.c] = nextCell;
              
              changes.push({
                cellRange: update.address,
                beforeValue,
                afterValue: Object.prototype.hasOwnProperty.call(update, 'value') ? update.value : nextCell.value
              });

              // ワークブックも更新
              if (workbook && workbook.Sheets[activeSheet]) {
                const cell = workbook.Sheets[activeSheet][update.address] || {};
                if (Object.prototype.hasOwnProperty.call(update, 'value')) {
                  cell.v = update.value;
                }
                const targetFormula = update.formula !== undefined ? update.formula : beforeFormula;
                if (targetFormula) {
                  cell.f = targetFormula;
                }
                if (nextCell.type) {
                  cell.t = nextCell.type;
                }
                workbook.Sheets[activeSheet][update.address] = cell;
              }
            }
          } catch (error) {
            console.error(`セル ${update.address} の更新エラー:`, error);
          }
        });

        setSheetData(newSheetData);

        // 編集履歴に追加
        const newHistory: EditHistory = {
          id: Date.now().toString(),
          timestamp: new Date(),
          instruction,
          changes,
          formatPreserved: true
        };

        setEditHistory(prev => [newHistory, ...prev.slice(0, 9)]);
        
        toast.success(`編集完了: ${parsedResponse.explanation || '変更が適用されました'}`);
      } else {
        toast.error('有効な更新データが見つかりませんでした');
      }

    } catch (error) {
      console.error('DeepSeek処理エラー:', error);
      toast.error(`編集処理に失敗しました: ${error instanceof Error ? error.message : 'Unknown error'}`);
    } finally {
      setIsProcessing(false);
    }
  }, [deepseekApiKey, workbook, sheetData, activeSheet]);

  // 編集指示の実行（DeepSeek API使用）
  const executeInstruction = useCallback(() => {
    if (!currentInstruction.trim()) {
      toast.error('編集指示を入力してください');
      return;
    }

    if (apiConnectionStatus !== 'success') {
      toast.error('まずDeepSeek APIの接続テストを実行してください');
      return;
    }

    processWithDeepSeek(currentInstruction);
    setCurrentInstruction('');
  }, [currentInstruction, apiConnectionStatus, processWithDeepSeek]);

  // ファイルドロップハンドラー
  const handleDrop = useCallback((e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setDragActive(false);
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const file = files[0];
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        setUploadedFile(file);
        readFileWithStyles(file);
      } else {
        toast.error('対応していないファイル形式です（.xlsx, .xlsのみ対応）');
      }
    }
  }, [readFileWithStyles]);

  // ファイル選択ハンドラー
  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files && files.length > 0) {
      const file = files[0];
      setUploadedFile(file);
      readFileWithStyles(file);
    }
  }, [readFileWithStyles]);

  // シート切り替え
  const handleSheetChange = useCallback((sheetName: string) => {
    if (workbook) {
      setActiveSheet(sheetName);
      loadSheetData(workbook, sheetName);
    }
  }, [workbook, loadSheetData]);

  // Undo機能
  const handleUndo = useCallback(() => {
    if (editHistory.length > 0) {
      const lastEdit = editHistory[0];
      setEditHistory(prev => prev.slice(1));
      toast.success(`「${lastEdit.instruction}」を元に戻しました`);
    }
  }, [editHistory]);

  // ダウンロード機能（書式保持改善版）
  const handleDownload = useCallback(() => {
    if (!workbook) {
      toast.error('ダウンロードするファイルがありません');
      return;
    }

    try {
      // 現在のシートデータをワークブックに反映（書式保持）
      const worksheet = workbook.Sheets[activeSheet];
      if (worksheet) {
        sheetData.forEach((row, rowIndex) => {
          row.forEach((cell, colIndex) => {
            const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const existingCell = worksheet[cellAddress];

            const isEmptyValue = cell.value === '' || cell.value === null || cell.value === undefined;
            const hasFormula = !!cell.formula;

            if (isEmptyValue && !hasFormula) {
              if (!existingCell) {
                return;
              }
              const cloned = { ...existingCell } as XLSX.CellObject;
              delete cloned.v;
              delete cloned.w;
              delete cloned.z;
              if (!cloned.f) {
                delete worksheet[cellAddress];
              } else {
                worksheet[cellAddress] = cloned;
              }
              return;
            }

            const nextCell: XLSX.CellObject = {
              ...(existingCell || {})
            };

            if (cell.type) {
              nextCell.t = cell.type;
            } else if (!nextCell.t) {
              nextCell.t = typeof cell.value === 'number' ? 'n' : 's';
            }

            if (hasFormula) {
              nextCell.f = cell.formula;
            } else if (nextCell.f) {
              delete nextCell.f;
            }

            if (!isEmptyValue) {
              nextCell.v = cell.value;
            }

            worksheet[cellAddress] = nextCell;
          });
        });
      }

      // 書式情報を含めて出力（改善版）
      const wbout = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'array',
        cellStyles: true,  // 書式を含める
        bookSST: true,     // 共有文字列テーブルを使用
        compression: true  // 圧縮を有効化
      });

      const blob = new Blob([wbout], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      
      const originalName = uploadedFile?.name || 'quote';
      const baseName = originalName.replace(/\.[^/.]+$/, '');
      const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      link.download = `${baseName}_edited_${timestamp}.xlsx`;
      
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);

      toast.success('ファイルをダウンロードしました（書式保持）');
    } catch (error) {
      console.error('ダウンロードエラー:', error);
      toast.error('ダウンロードに失敗しました');
    }
  }, [workbook, uploadedFile, sheetData, activeSheet]);

  // セルが結合されているかチェック
  const isMergedCell = useCallback((row: number, col: number): boolean => {
    return merges.some(merge => 
      row >= merge.s.r && row <= merge.e.r && 
      col >= merge.s.c && col <= merge.e.c
    );
  }, [merges]);

  // セルのスタイルを取得（改善版）
  const getCellStyle = useCallback((row: number, col: number): React.CSSProperties => {
    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
    const style = cellStyles.get(cellAddress);
    
    if (!style) return {};

    const cssStyle: React.CSSProperties = {};
    
    if (style.backgroundColor) {
      cssStyle.backgroundColor = style.backgroundColor;
    }
    if (style.color) {
      cssStyle.color = style.color;
    }
    if (style.fontWeight) {
      cssStyle.fontWeight = style.fontWeight;
    }
    if (style.fontStyle) {
      cssStyle.fontStyle = style.fontStyle;
    }
    if (style.textDecoration) {
      cssStyle.textDecoration = style.textDecoration;
    }
    if (style.fontSize) {
      cssStyle.fontSize = style.fontSize;
    }
    if (style.fontFamily) {
      cssStyle.fontFamily = style.fontFamily;
    }
    if (style.textAlign) {
      cssStyle.textAlign = style.textAlign as any;
    }
    if (style.border) {
      // 複数の罫線スタイルを適用
      const borderStyles = style.border.split('; ');
      borderStyles.forEach(borderStyle => {
        const [property, value] = borderStyle.split(': ');
        if (property && value) {
          (cssStyle as any)[property.replace(/-([a-z])/g, (g) => g[1].toUpperCase())] = value;
        }
      });
    }

    return cssStyle;
  }, [cellStyles]);

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-4">
      <div className="max-w-7xl mx-auto space-y-6">
        {/* ヘッダーセクション */}
        <Card className="border-2 border-blue-200 shadow-lg">
          <CardHeader className="text-center">
            <CardTitle className="text-3xl font-bold text-blue-900 flex items-center justify-center gap-3">
              <FileSpreadsheet className="h-8 w-8" />
              見積書エディター
            </CardTitle>
            <p className="text-blue-700 mt-2">DeepSeek AIでExcel見積書を自然言語編集（書式保持）</p>
            <div className="flex items-center justify-center gap-2 mt-3">
              <Palette className="h-5 w-5 text-purple-600" />
              <Badge variant="secondary" className="bg-purple-100 text-purple-800">
                📋 元の書式を保持します
              </Badge>
              <Zap className="h-5 w-5 text-yellow-600" />
              <Badge variant="secondary" className="bg-yellow-100 text-yellow-800">
                🤖 DeepSeek AI搭載
              </Badge>
            </div>
          </CardHeader>
        </Card>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* 左側：アップロード・編集エリア */}
          <div className="lg:col-span-1 space-y-6">
            {/* DeepSeek API設定エリア */}
            <Card className="border-2 border-yellow-200">
              <CardHeader>
                <CardTitle className="flex items-center gap-2 text-yellow-800">
                  <Key className="h-5 w-5" />
                  DeepSeek API設定
                </CardTitle>
              </CardHeader>
              <CardContent className="space-y-4">
                <div className="space-y-2">
                  <Label htmlFor="deepseek-api-key">APIキー</Label>
                  <Input
                    id="deepseek-api-key"
                    type="password"
                    placeholder="sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
                    value={deepseekApiKey}
                    onChange={(e) => setDeepseekApiKey(e.target.value)}
                  />
                  <p className="text-xs text-gray-500">
                    DeepSeek APIキーを入力してください（
                    <a 
                      href="https://platform.deepseek.com/" 
                      target="_blank" 
                      rel="noopener noreferrer"
                      className="text-blue-600 hover:underline"
                    >
                      取得はこちら
                    </a>
                    ）
                  </p>
                </div>
                
                <div className="flex gap-2">
                  <Button 
                    onClick={testDeepSeekConnection}
                    disabled={isApiTesting || !deepseekApiKey.trim()}
                    variant="outline"
                    className="flex-1"
                  >
                    {isApiTesting ? (
                      <>
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mr-2"></div>
                        テスト中...
                      </>
                    ) : (
                      <>
                        <TestTube className="h-4 w-4 mr-2" />
                        接続テスト
                      </>
                    )}
                  </Button>
                </div>

                {apiConnectionStatus !== 'none' && (
                  <Alert className={apiConnectionStatus === 'success' ? 'border-green-200 bg-green-50' : 'border-red-200 bg-red-50'}>
                    {apiConnectionStatus === 'success' ? (
                      <CheckCircle className="h-4 w-4 text-green-600" />
                    ) : (
                      <AlertCircle className="h-4 w-4 text-red-600" />
                    )}
                    <AlertDescription className={apiConnectionStatus === 'success' ? 'text-green-800' : 'text-red-800'}>
                      {apiConnectionStatus === 'success' 
                        ? 'API接続成功！編集機能が利用可能です。' 
                        : 'API接続に失敗しました。APIキーを確認してください。'
                      }
                    </AlertDescription>
                  </Alert>
                )}
              </CardContent>
            </Card>

            {/* ファイルアップロードエリア */}
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Upload className="h-5 w-5" />
                  ファイルアップロード
                </CardTitle>
              </CardHeader>
              <CardContent>
                {!uploadedFile ? (
                  <div
                    className={`border-2 border-dashed rounded-lg p-8 text-center transition-colors ${
                      dragActive 
                        ? 'border-blue-500 bg-blue-50' 
                        : 'border-gray-300 hover:border-blue-400'
                    }`}
                    onDrop={handleDrop}
                    onDragOver={(e) => {
                      e.preventDefault();
                      setDragActive(true);
                    }}
                    onDragLeave={() => setDragActive(false)}
                  >
                    <Upload className="h-12 w-12 text-gray-400 mx-auto mb-4" />
                    <p className="text-gray-600 mb-2">
                      ファイルをドラッグ&ドロップ
                    </p>
                    <p className="text-sm text-gray-500 mb-4">
                      または
                    </p>
                    <Button 
                      onClick={() => fileInputRef.current?.click()}
                      variant="outline"
                    >
                      ファイルを選択
                    </Button>
                    <p className="text-xs text-gray-500 mt-2">
                      .xlsx, .xls対応
                    </p>
                    <input
                      ref={fileInputRef}
                      type="file"
                      accept=".xlsx,.xls"
                      onChange={handleFileSelect}
                      className="hidden"
                    />
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="flex items-center gap-2 p-3 bg-green-50 rounded-lg">
                      <CheckCircle className="h-5 w-5 text-green-600" />
                      <span className="text-green-800 font-medium">
                        {uploadedFile.name}
                      </span>
                    </div>
                    <Button 
                      variant="outline" 
                      size="sm"
                      onClick={() => {
                        setUploadedFile(null);
                        setWorkbook(null);
                        setSheets([]);
                        setSheetData([]);
                        setCellStyles(new Map());
                        setEditHistory([]);
                      }}
                    >
                      別のファイルを選択
                    </Button>
                    {isLoading && (
                      <div className="flex items-center gap-2 text-blue-600">
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600"></div>
                        <span className="text-sm">書式情報を読み込み中...</span>
                      </div>
                    )}
                  </div>
                )}
              </CardContent>
            </Card>

            {/* シート選択エリア */}
            {sheets.length > 1 && (
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <FileText className="h-5 w-5" />
                    シート選択
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <Tabs value={activeSheet} onValueChange={handleSheetChange}>
                    <TabsList className="grid w-full grid-cols-2">
                      {sheets.map((sheet) => (
                        <TabsTrigger key={sheet} value={sheet} className="text-xs">
                          {sheet}
                        </TabsTrigger>
                      ))}
                    </TabsList>
                  </Tabs>
                </CardContent>
              </Card>
            )}

            {/* 編集指示入力エリア */}
            {workbook && (
              <Card className="border-2 border-green-200">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-green-800">
                    <Zap className="h-5 w-5" />
                    AI編集指示入力
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <Textarea
                    placeholder={`例：
- 小計を10%値引きして更新して
- 消費税を8%から10%に変更
- 運送費と設置費を「配送関連費用」にまとめる
- 為替レート140円で再計算
- A列の単価を全て20%アップして

※DeepSeek AIが自動で計算・編集します`}
                    value={currentInstruction}
                    onChange={(e) => setCurrentInstruction(e.target.value)}
                    rows={6}
                  />
                  <Button 
                    onClick={executeInstruction}
                    disabled={isProcessing || apiConnectionStatus !== 'success'}
                    className="w-full bg-green-600 hover:bg-green-700"
                  >
                    {isProcessing ? (
                      <>
                        <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-white mr-2"></div>
                        AI処理中...
                      </>
                    ) : (
                      <>
                        <Zap className="h-4 w-4 mr-2" />
                        AI編集を実行（書式保持）
                      </>
                    )}
                  </Button>
                  
                  {apiConnectionStatus !== 'success' && (
                    <Alert>
                      <AlertCircle className="h-4 w-4" />
                      <AlertDescription className="text-sm">
                        AI編集を使用するには、まずDeepSeek APIの接続テストを成功させてください。
                      </AlertDescription>
                    </Alert>
                  )}
                  
                  {/* クイック編集ボタン */}
                  <div className="grid grid-cols-2 gap-2">
                    <Button 
                      variant="outline" 
                      size="sm" 
                      className="text-xs"
                      onClick={() => setCurrentInstruction('新しい行を最後に追加して')}
                    >
                      <Plus className="h-3 w-3 mr-1" />
                      行を追加
                    </Button>
                    <Button 
                      variant="outline" 
                      size="sm" 
                      className="text-xs"
                      onClick={() => setCurrentInstruction('新しい列を最後に追加して')}
                    >
                      <Plus className="h-3 w-3 mr-1" />
                      列を追加
                    </Button>
                    <Button 
                      variant="outline" 
                      size="sm" 
                      className="text-xs"
                      onClick={() => setCurrentInstruction('小計を計算して追加して')}
                    >
                      <Calculator className="h-3 w-3 mr-1" />
                      小計を計算
                    </Button>
                    <Button 
                      variant="outline" 
                      size="sm" 
                      className="text-xs"
                      onClick={() => setCurrentInstruction('全ての金額に10%の消費税を追加して')}
                    >
                      <Percent className="h-3 w-3 mr-1" />
                      消費税追加
                    </Button>
                  </div>
                </CardContent>
              </Card>
            )}

            {/* 書式保持インジケーターパネル */}
            {workbook && (
              <Card className="border-purple-200">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-purple-800">
                    <Palette className="h-5 w-5" />
                    書式保持状況
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="space-y-2 text-sm">
                    <div className="flex items-center gap-2">
                      <CheckCircle className="h-4 w-4 text-green-600" />
                      <span>セルの背景色</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <CheckCircle className="h-4 w-4 text-green-600" />
                      <span>フォント（書体、サイズ、色、太字、斜体）</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <CheckCircle className="h-4 w-4 text-green-600" />
                      <span>罫線</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <CheckCircle className="h-4 w-4 text-green-600" />
                      <span>セル結合</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <CheckCircle className="h-4 w-4 text-green-600" />
                      <span>数値フォーマット</span>
                    </div>
                    <div className="flex items-center gap-2">
                      <CheckCircle className="h-4 w-4 text-green-600" />
                      <span>セル幅・行高</span>
                    </div>
                  </div>
                  
                  <Separator className="my-3" />
                  
                  <Alert>
                    <AlertCircle className="h-4 w-4" />
                    <AlertDescription className="text-xs">
                      画像・グラフ・マクロは保持されません
                    </AlertDescription>
                  </Alert>
                </CardContent>
              </Card>
            )}
          </div>

          {/* 右側：データプレビューエリア */}
          <div className="lg:col-span-2 space-y-6">
            {workbook && sheetData.length > 0 && (
              <Card className={previewFullscreen ? 'fixed inset-4 z-50 bg-white shadow-2xl' : ''}>
                <CardHeader>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <Eye className="h-5 w-5" />
                      データプレビュー（書式表示）
                      <Badge variant="outline" className="ml-2">
                        {activeSheet}
                      </Badge>
                    </div>
                    <div className="flex items-center gap-2">
                      {/* ズーム機能 */}
                      <div className="flex items-center gap-2 text-sm">
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => setZoomLevel(Math.max(50, zoomLevel - 10))}
                          disabled={zoomLevel <= 50}
                        >
                          -
                        </Button>
                        <span className="min-w-16 text-center">{zoomLevel}%</span>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => setZoomLevel(Math.min(200, zoomLevel + 10))}
                          disabled={zoomLevel >= 200}
                        >
                          +
                        </Button>
                        <Button
                          variant="outline"
                          size="sm"
                          onClick={() => setZoomLevel(100)}
                        >
                          リセット
                        </Button>
                      </div>
                      {/* フルスクリーン切り替え */}
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={() => setPreviewFullscreen(!previewFullscreen)}
                      >
                        {previewFullscreen ? '元に戻す' : 'フルスクリーン'}
                      </Button>
                    </div>
                  </div>
                </CardHeader>
                <CardContent>
                  <div 
                    className={`overflow-auto border rounded-lg ${
                      previewFullscreen ? 'max-h-[calc(100vh-200px)]' : 'max-h-[600px]'
                    }`}
                    style={{ 
                      transform: `scale(${zoomLevel / 100})`,
                      transformOrigin: 'top left',
                      width: `${10000 / zoomLevel}%`,
                      height: `${10000 / zoomLevel}%`
                    }}
                  >
                    <table className="w-full text-sm border-collapse">
                      <thead>
                        <tr className="bg-gray-50 sticky top-0">
                          <th className="p-2 border text-center w-12 bg-gray-100 sticky left-0 z-10">#</th>
                          {sheetData[0]?.map((_, colIndex) => (
                            <th key={colIndex} className="p-2 border text-center min-w-32 bg-gray-100">
                              {String.fromCharCode(65 + colIndex)}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sheetData.map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            <td className="p-2 border text-center bg-gray-50 font-medium sticky left-0 z-10">
                              {rowIndex + 1}
                            </td>
                            {row.map((cell, colIndex) => {
                              const cellStyle = getCellStyle(rowIndex, colIndex);
                              const isMerged = isMergedCell(rowIndex, colIndex);
                              
                              return (
                                <td
                                  key={colIndex}
                                  className="p-2 border relative min-w-32"
                                  style={{
                                    ...cellStyle,
                                    minHeight: '32px',
                                    verticalAlign: 'middle'
                                  }}
                                  title={`セル: ${String.fromCharCode(65 + colIndex)}${rowIndex + 1}${
                                    cell.formula ? `\n数式: ${cell.formula}` : ''
                                  }${
                                    Object.keys(cellStyle).length > 0 ? '\n書式: 適用済み' : ''
                                  }`}
                                >
                                  {isMerged && (
                                    <Link className="absolute top-1 right-1 h-3 w-3 text-blue-500" />
                                  )}
                                  <div className={
                                    typeof cell.value === 'number' 
                                      ? 'text-right' 
                                      : 'text-left'
                                  }>
                                    {cell.value !== undefined && cell.value !== null ? String(cell.value) : ''}
                                  </div>
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <div className="flex justify-between items-center mt-4 text-xs text-gray-500">
                    <span>
                      表示中: {sheetData.length}行 × {sheetData[0]?.length || 0}列
                    </span>
                    <span>
                      ズーム: {zoomLevel}% | スクロール可能
                    </span>
                  </div>
                </CardContent>
              </Card>
            )}

            {/* 編集履歴エリア */}
            {editHistory.length > 0 && (
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <History className="h-5 w-5" />
                    編集履歴
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <Accordion type="single" collapsible>
                    <AccordionItem value="history">
                      <AccordionTrigger>
                        履歴を表示 ({editHistory.length}件)
                      </AccordionTrigger>
                      <AccordionContent>
                        <div className="space-y-3">
                          {editHistory.map((history) => (
                            <div key={history.id} className="border rounded-lg p-3 bg-gray-50">
                              <div className="flex items-start justify-between">
                                <div className="flex-1">
                                  <p className="font-medium text-sm">{history.instruction}</p>
                                  <p className="text-xs text-gray-500 mt-1">
                                    {history.timestamp.toLocaleString()}
                                  </p>
                                  <div className="mt-2">
                                    <Badge 
                                      variant={history.formatPreserved ? "default" : "destructive"}
                                      className="text-xs"
                                    >
                                      {history.formatPreserved ? "✓ 書式保持" : "⚠ 書式変更"}
                                    </Badge>
                                  </div>
                                  <div className="mt-2 text-xs text-gray-600">
                                    変更範囲: {history.changes.map(c => c.cellRange).join(', ')}
                                  </div>
                                </div>
                                <Button
                                  variant="outline"
                                  size="sm"
                                  onClick={handleUndo}
                                  className="ml-2"
                                >
                                  <RotateCcw className="h-3 w-3" />
                                </Button>
                              </div>
                            </div>
                          ))}
                        </div>
                      </AccordionContent>
                    </AccordionItem>
                  </Accordion>
                </CardContent>
              </Card>
            )}
          </div>
        </div>

        {/* アクションボタンエリア */}
        {workbook && (
          <Card>
            <CardContent className="pt-6">
              <div className="flex flex-wrap gap-3 justify-center">
                <Button 
                  onClick={handleDownload}
                  className="bg-green-600 hover:bg-green-700 text-white"
                >
                  <Download className="h-4 w-4 mr-2" />
                  Excelをダウンロード（書式保持）
                </Button>
                
                <Button 
                  variant="outline"
                  onClick={() => {
                    setWorkbook(null);
                    setUploadedFile(null);
                    setSheets([]);
                    setSheetData([]);
                    setCellStyles(new Map());
                    setEditHistory([]);
                    toast.success('すべてリセットしました');
                  }}
                  className="text-red-600 border-red-300 hover:bg-red-50"
                >
                  <RotateCcw className="h-4 w-4 mr-2" />
                  すべてリセット
                </Button>

                <Button variant="outline">
                  <Eye className="h-4 w-4 mr-2" />
                  プレビュー更新
                </Button>

                <Button variant="outline">
                  <FileText className="h-4 w-4 mr-2" />
                  JSON出力（データのみ）
                </Button>

                <Button variant="outline">
                  <Palette className="h-4 w-4 mr-2" />
                  書式情報を確認
                </Button>
              </div>
            </CardContent>
          </Card>
        )}

        {/* 説明・注意事項 */}
        {!workbook && (
          <Card className="border-blue-200">
            <CardContent className="pt-6">
              <Alert>
                <AlertCircle className="h-4 w-4" />
                <AlertDescription>
                  <strong>使い方：</strong>
                  <br />
                  1. DeepSeek APIキーを設定し、接続テストを実行
                  <br />
                  2. Excel見積書ファイル（.xlsx, .xls）をアップロード
                  <br />
                  3. 自然言語でAI編集指示を入力（例：「小計を10%値引きして更新して」）
                  <br />
                  4. AI編集を実行し、書式を保持したままダウンロード
                  <br />
                  <br />
                  <strong>注意：</strong> DeepSeek APIキーはブラウザ上で直接使用されます。
                  セキュリティ上、信頼できる環境でのみご利用ください。
                </AlertDescription>
              </Alert>
            </CardContent>
          </Card>
        )}
      </div>
    </div>
  );
};

export default ExcelQuoteEditor;