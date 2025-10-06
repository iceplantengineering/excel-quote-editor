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
  ChevronUp
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Textarea } from '@/components/ui/textarea';
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

          // スタイル情報を抽出
          if (cell.s) {
            const style = cell.s;
            
            // 背景色
            if (style.fill && style.fill.fgColor) {
              const color = style.fill.fgColor;
              if (color.rgb) {
                cellStyle.backgroundColor = `#${color.rgb}`;
              }
            }

            // フォント情報
            if (style.font) {
              const font = style.font;
              if (font.color && font.color.rgb) {
                cellStyle.color = `#${font.color.rgb}`;
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

            // 罫線（簡易実装）
            if (style.border) {
              cellStyle.border = '1px solid #ccc';
            }
          }

          cellData.style = cellStyle;
          styles.set(cellAddress, cellStyle);
        }

        row.push(cellData);
      }
      data.push(row);
    }

    setSheetData(data);
    setCellStyles(styles);
  }, []);

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

  // 編集指示の実行（シミュレーション）
  const executeInstruction = useCallback(() => {
    if (!currentInstruction.trim()) {
      toast.error('編集指示を入力してください');
      return;
    }

    // シミュレーション用の編集処理
    const newHistory: EditHistory = {
      id: Date.now().toString(),
      timestamp: new Date(),
      instruction: currentInstruction,
      changes: [
        {
          cellRange: 'A1:B2',
          beforeValue: '変更前の値',
          afterValue: '変更後の値'
        }
      ],
      formatPreserved: true
    };

    setEditHistory(prev => [newHistory, ...prev.slice(0, 9)]); // 最新10件を保持
    setCurrentInstruction('');
    
    toast.success('編集指示を実行しました（デモ）');
  }, [currentInstruction]);

  // Undo機能
  const handleUndo = useCallback(() => {
    if (editHistory.length > 0) {
      const lastEdit = editHistory[0];
      setEditHistory(prev => prev.slice(1));
      toast.success(`「${lastEdit.instruction}」を元に戻しました`);
    }
  }, [editHistory]);

  // ダウンロード機能（書式保持）
  const handleDownload = useCallback(() => {
    if (!workbook) {
      toast.error('ダウンロードするファイルがありません');
      return;
    }

    try {
      // 書式情報を含めて出力
      const wbout = XLSX.write(workbook, {
        bookType: 'xlsx',
        type: 'array',
        cellStyles: true,  // 書式を含める
        bookSST: false
      });

      const blob = new Blob([wbout], { type: 'application/octet-stream' });
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
  }, [workbook, uploadedFile]);

  // セルが結合されているかチェック
  const isMergedCell = useCallback((row: number, col: number): boolean => {
    return merges.some(merge => 
      row >= merge.s.r && row <= merge.e.r && 
      col >= merge.s.c && col <= merge.e.c
    );
  }, [merges]);

  // セルのスタイルを取得
  const getCellStyle = useCallback((row: number, col: number): React.CSSProperties => {
    const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
    const style = cellStyles.get(cellAddress);
    
    if (!style) return {};

    return {
      backgroundColor: style.backgroundColor,
      color: style.color,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      textDecoration: style.textDecoration,
      fontSize: style.fontSize,
      fontFamily: style.fontFamily,
      textAlign: style.textAlign as any,
      border: style.border,
    };
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
            <p className="text-blue-700 mt-2">自然言語でExcel見積書を編集（書式保持）</p>
            <div className="flex items-center justify-center gap-2 mt-3">
              <Palette className="h-5 w-5 text-purple-600" />
              <Badge variant="secondary" className="bg-purple-100 text-purple-800">
                📋 元の書式を保持します
              </Badge>
            </div>
          </CardHeader>
        </Card>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* 左側：アップロード・編集エリア */}
          <div className="lg:col-span-1 space-y-6">
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
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <Calculator className="h-5 w-5" />
                    編集指示入力
                  </CardTitle>
                </CardHeader>
                <CardContent className="space-y-4">
                  <Textarea
                    placeholder={`例：
- 小計を10%値引きして更新して
- 消費税を8%から10%に変更
- 運送費と設置費を「配送関連費用」にまとめる
- 為替レート140円で再計算

※書式は自動的に保持されます`}
                    value={currentInstruction}
                    onChange={(e) => setCurrentInstruction(e.target.value)}
                    rows={6}
                  />
                  <Button 
                    onClick={executeInstruction}
                    className="w-full bg-blue-600 hover:bg-blue-700"
                  >
                    編集を実行（書式保持）
                  </Button>
                  
                  {/* クイック編集ボタン */}
                  <div className="grid grid-cols-2 gap-2">
                    <Button variant="outline" size="sm" className="text-xs">
                      <Plus className="h-3 w-3 mr-1" />
                      列を追加
                      <Badge variant="secondary" className="ml-1 text-xs">書式継承</Badge>
                    </Button>
                    <Button variant="outline" size="sm" className="text-xs">
                      <Plus className="h-3 w-3 mr-1" />
                      行を追加
                      <Badge variant="secondary" className="ml-1 text-xs">書式継承</Badge>
                    </Button>
                    <Button variant="outline" size="sm" className="text-xs">
                      <Calculator className="h-3 w-3 mr-1" />
                      小計を計算
                      <Badge variant="secondary" className="ml-1 text-xs">書式継承</Badge>
                    </Button>
                    <Button variant="outline" size="sm" className="text-xs">
                      <Percent className="h-3 w-3 mr-1" />
                      パーセント追加
                      <Badge variant="secondary" className="ml-1 text-xs">書式継承</Badge>
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
              <Card>
                <CardHeader>
                  <CardTitle className="flex items-center gap-2">
                    <Eye className="h-5 w-5" />
                    データプレビュー（書式表示）
                    <Badge variant="outline" className="ml-2">
                      {activeSheet}
                    </Badge>
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="overflow-auto max-h-96 border rounded-lg">
                    <table className="w-full text-sm">
                      <thead>
                        <tr className="bg-gray-50">
                          <th className="p-2 border text-center w-12">#</th>
                          {sheetData[0]?.map((_, colIndex) => (
                            <th key={colIndex} className="p-2 border text-center min-w-24">
                              {String.fromCharCode(65 + colIndex)}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {sheetData.slice(0, 50).map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            <td className="p-2 border text-center bg-gray-50 font-medium">
                              {rowIndex + 1}
                            </td>
                            {row.map((cell, colIndex) => (
                              <td
                                key={colIndex}
                                className="p-2 border relative"
                                style={getCellStyle(rowIndex, colIndex)}
                                title={`セル: ${String.fromCharCode(65 + colIndex)}${rowIndex + 1}${
                                  cell.formula ? `\n数式: ${cell.formula}` : ''
                                }`}
                              >
                                {isMergedCell(rowIndex, colIndex) && (
                                  <Link className="absolute top-1 right-1 h-3 w-3 text-blue-500" />
                                )}
                                <span className={
                                  typeof cell.value === 'number' 
                                    ? 'text-right block' 
                                    : 'text-left block'
                                }>
                                  {cell.value || ''}
                                </span>
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  {sheetData.length > 50 && (
                    <p className="text-xs text-gray-500 mt-2">
                      ※ パフォーマンス向上のため、最初の50行のみ表示しています
                    </p>
                  )}
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
                  1. Excel見積書ファイル（.xlsx, .xls）をアップロード
                  <br />
                  2. 自然言語で編集指示を入力（例：「小計を10%値引きして更新して」）
                  <br />
                  3. 編集を実行し、書式を保持したままダウンロード
                  <br />
                  <br />
                  <strong>注意：</strong> 現在はデモ版のため、実際の編集処理はシミュレーションです。
                  実際の自然言語処理はClaudeとの対話で実現されます。
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