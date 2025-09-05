import {
  GoogleGenerativeAI,
} from '@google/generative-ai';
import { google } from 'googleapis';
import * as dotenv from 'dotenv';

dotenv.config();

interface BusinessCardData {
  fileId: string;
  name?: string;
  furigana?: string;
  companyName?: string;
  companyFurigana?: string;
  department?: string;
  position?: string;
  postalCode?: string;
  address?: string;
  email?: string;
  companyPhone?: string;
  personalPhone?: string;
  fax?: string;
}

class BusinessCardProcessor {
  private ai: GoogleGenerativeAI;
  private sheets: any;
  private drive: any;
  private model: any;
  
  private readonly BUSINESS_CARD_FOLDER_ID = '1Q29bAIoQ__PADA2NefymTpsOO_yAg9ee';
  private readonly PROCESSED_FOLDER_ID = '16LAj4yAM2cyk-tUlY7WYFTkbmB5Krvjx';
  private readonly SHEET_NAME = '名刺情報';
  
  private readonly LEGAL_FORMS = [
    // 漢字表記
    { pattern: '株式会社', furigana: 'カブシキガイシャ' },
    { pattern: '有限会社', furigana: 'ユウゲンガイシャ' },
    { pattern: '合同会社', furigana: 'ゴウドウガイシャ' },
    { pattern: '合資会社', furigana: 'ゴウシガイシャ' },
    { pattern: '合名会社', furigana: 'ゴウメイガイシャ' },
    { pattern: '一般社団法人', furigana: 'イッパンシャダンホウジン' },
    { pattern: '公益社団法人', furigana: 'コウエキシャダンホウジン' },
    { pattern: '一般財団法人', furigana: 'イッパンザイダンホウジン' },
    { pattern: '公益財団法人', furigana: 'コウエキザイダンホウジン' },
    { pattern: '特定非営利活動法人', furigana: 'トクテイヒエイリカツドウホウジン' },
    { pattern: 'NPO法人', furigana: 'エヌピーオーホウジン' },
    // 省略表記
    { pattern: '（株）', furigana: 'カブシキガイシャ' },
    { pattern: '(株)', furigana: 'カブシキガイシャ' },
    { pattern: '（有）', furigana: 'ユウゲンガイシャ' },
    { pattern: '(有)', furigana: 'ユウゲンガイシャ' },
    { pattern: '（同）', furigana: 'ゴウドウガイシャ' },
    { pattern: '(同)', furigana: 'ゴウドウガイシャ' },
    { pattern: '（資）', furigana: 'ゴウシガイシャ' },
    { pattern: '(資)', furigana: 'ゴウシガイシャ' },
    { pattern: '（名）', furigana: 'ゴウメイガイシャ' },
    { pattern: '(名)', furigana: 'ゴウメイガイシャ' },
    // 英語表記
    { pattern: 'Corporation', furigana: 'コーポレーション' },
    { pattern: 'Corp.', furigana: 'コーポレーション' },
    { pattern: 'Company', furigana: 'カンパニー' },
    { pattern: 'Co.', furigana: 'カンパニー' },
    { pattern: 'Limited', furigana: 'リミテッド' },
    { pattern: 'Ltd.', furigana: 'リミテッド' },
    { pattern: 'Inc.', furigana: 'インク' },
    { pattern: 'LLC', furigana: 'エルエルシー' }
  ];

  constructor() {
    this.ai = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || '');
    this.model = this.ai.getGenerativeModel({ model: 'gemini-2.5-pro' });
    
    const auth = new google.auth.GoogleAuth({
      keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_KEY_FILE,
      scopes: [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'
      ]
    });
    
    this.sheets = google.sheets({ version: 'v4', auth });
    this.drive = google.drive({ version: 'v3', auth });
  }

  private normalizeCompanyName(companyName: string): string {
    // 会社名を正規化（重複チェック用）
    return companyName
      .replace(/[　 ]+/g, '') // 全角・半角スペースを除去
      .replace(/[・・·]/g, '') // 中点を除去
      .replace(/[（）()]/g, '') // カッコを除去
      .replace(/株式会社|有限会社|合同会社|合資会社|合名会社/g, '') // 法人格を除去
      .replace(/（株）|（有）|（同）|（資）|（名）/g, '') // 省略法人格を除去
      .replace(/\(株\)|\(有\)|\(同\)|\(資\)|\(名\)/g, '') // 省略法人格（半角カッコ）を除去
      .toUpperCase() // 大文字に統一
      .trim();
  }

  private removeCompanyLegalForm(companyName: string | undefined, companyFurigana: string | undefined): string | undefined {
    if (!companyFurigana || !companyName) return companyFurigana;
    
    // 漢字の会社名から法人格を検出
    let detectedLegalForm = null;
    let legalFormPosition = 'none'; // 'prefix', 'suffix', 'none'
    
    for (const legalForm of this.LEGAL_FORMS) {
      // 前方にあるかチェック
      if (companyName.startsWith(legalForm.pattern)) {
        detectedLegalForm = legalForm;
        legalFormPosition = 'prefix';
        break;
      }
      // 後方にあるかチェック
      if (companyName.endsWith(legalForm.pattern)) {
        detectedLegalForm = legalForm;
        legalFormPosition = 'suffix';
        break;
      }
    }
    
    if (!detectedLegalForm) {
      return companyFurigana;
    }
    
    let result = companyFurigana;
    
    // 検出された法人格のフリガナを除去
    if (legalFormPosition === 'prefix') {
      // 前方から除去
      const variations = [
        detectedLegalForm.furigana,
        detectedLegalForm.furigana.replace(/ガイシャ/, 'カイシャ')
      ];
      
      for (const variation of variations) {
        if (result.startsWith(variation)) {
          result = result.slice(variation.length);
          break;
        }
      }
    } else if (legalFormPosition === 'suffix') {
      // 後方から除去
      const variations = [
        detectedLegalForm.furigana,
        detectedLegalForm.furigana.replace(/ガイシャ/, 'カイシャ')
      ];
      
      for (const variation of variations) {
        if (result.endsWith(variation)) {
          result = result.slice(0, -variation.length);
          break;
        }
      }
    }
    
    // スペース、中点、カッコの調整
    result = result.replace(/^[\s・・（）()]+|[\s・・（）()]+$/g, '').trim();
    
    return result;
  }

  async initializeSheet(spreadsheetId: string): Promise<void> {
    try {
      // まずシートの存在確認
      const sheetInfo = await this.sheets.spreadsheets.get({
        spreadsheetId,
        fields: 'sheets.properties'
      });
      
      const sheetExists = sheetInfo.data.sheets?.some(
        (sheet: any) => sheet.properties?.title === this.SHEET_NAME
      );

      if (!sheetExists) {
        console.log(`シート「${this.SHEET_NAME}」が見つかりません。新規作成します。`);
        
        // 新しいシートを作成
        await this.sheets.spreadsheets.batchUpdate({
          spreadsheetId,
          resource: {
            requests: [{
              addSheet: {
                properties: {
                  title: this.SHEET_NAME
                }
              }
            }]
          }
        });
        
        console.log(`シート「${this.SHEET_NAME}」を作成しました`);
      }

      // ヘッダーを設定
      const headers = [
        'ファイルid', '名前', 'フリガナ', '社名', '社名フリガナ', 
        '部署', '肩書', '郵便番号', '住所', 'メールアドレス',
        '会社電話番号', '個人電話番号', 'FAX'
      ];

      await this.sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `'${this.SHEET_NAME}'!A1:M1`,
        valueInputOption: 'RAW',
        resource: {
          values: [headers]
        }
      });

      console.log('スプレッドシートのヘッダーを作成しました');
    } catch (error) {
      console.error('ヘッダー作成エラー:', error);
      throw error;
    }
  }

  async getFilesInFolder(folderId: string): Promise<string[]> {
    try {
      const response = await this.drive.files.list({
        q: `'${folderId}' in parents and trashed=false and (mimeType contains 'image/')`,
        fields: 'files(id, name)'
      });

      return response.data.files?.map((file: any) => file.id) || [];
    } catch (error) {
      console.error('フォルダ内ファイル取得エラー:', error);
      throw error;
    }
  }

  async getExistingFileIds(spreadsheetId: string): Promise<string[]> {
    try {
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${this.SHEET_NAME}'!A:A`
      });

      const values = response.data.values || [];
      return values.slice(1).map((row: string[]) => row[0]).filter(Boolean);
    } catch (error) {
      console.error('既存ファイルID取得エラー:', error);
      return [];
    }
  }

  async getExistingData(spreadsheetId: string): Promise<{name: string, companyName: string, originalCompanyName: string}[]> {
    try {
      const response = await this.sheets.spreadsheets.values.get({
        spreadsheetId,
        range: `'${this.SHEET_NAME}'!B:D`
      });

      const values = response.data.values || [];
      return values.slice(1).map((row: string[]) => {
        const originalCompanyName = (row[2] || '').trim();
        return {
          name: (row[0] || '').trim(),
          companyName: this.normalizeCompanyName(originalCompanyName),
          originalCompanyName: originalCompanyName
        };
      });
    } catch (error) {
      console.error('既存データ取得エラー:', error);
      return [];
    }
  }

  async isDuplicateData(spreadsheetId: string, data: BusinessCardData): Promise<boolean> {
    try {
      const existingData = await this.getExistingData(spreadsheetId);
      const newName = (data.name || '').trim();
      const newCompanyName = (data.companyName || '').trim();
      const normalizedNewCompanyName = this.normalizeCompanyName(newCompanyName);

      // 名前が空の場合は重複チェックをスキップ
      if (!newName) {
        return false;
      }

      // 名前の一致をチェック
      const nameMatch = existingData.find(existing => existing.name === newName);
      if (!nameMatch) {
        return false;
      }

      // 名前が一致した場合のみ会社名をチェック（正規化して比較）
      const isDuplicate = nameMatch.companyName === normalizedNewCompanyName;
      
      if (isDuplicate) {
        console.log(`重複検出: 名前="${newName}", 新会社名="${newCompanyName}" (正規化: "${normalizedNewCompanyName}"), 既存会社名="${nameMatch.originalCompanyName}" (正規化: "${nameMatch.companyName}")`);
      }
      
      return isDuplicate;
    } catch (error) {
      console.error('重複チェックエラー:', error);
      return false;
    }
  }

  async getNewFiles(spreadsheetId: string): Promise<string[]> {
    const folderFileIds = await this.getFilesInFolder(this.BUSINESS_CARD_FOLDER_ID);
    const existingFileIds = await this.getExistingFileIds(spreadsheetId);
    
    return folderFileIds.filter(fileId => !existingFileIds.includes(fileId));
  }

  async downloadImageAsBase64(fileId: string): Promise<string> {
    try {
      const response = await this.drive.files.get({
        fileId,
        alt: 'media'
      }, { responseType: 'arraybuffer' });

      return Buffer.from(response.data).toString('base64');
    } catch (error) {
      console.error('画像ダウンロードエラー:', error);
      throw error;
    }
  }

  async analyzeBusinessCard(fileId: string): Promise<BusinessCardData> {
    try {
      const imageBase64 = await this.downloadImageAsBase64(fileId);
      
      const image = {
        inlineData: {
          data: imageBase64,
          mimeType: 'image/jpeg'
        }
      };

      const prompt = `名刺画像から以下の情報を抽出してJSONで返してください：
- name: 名前
- furigana: 名前のフリガナ（カタカナで出力してください）
- companyName: 社名
- companyFurigana: 社名のフリガナ（カタカナで出力してください）
- department: 部署
- position: 肩書
- postalCode: 郵便番号
- address: 住所
- email: メールアドレス
- companyPhone: 会社電話番号（代表番号やTEL、会社名の近くにある番号）
- personalPhone: 個人電話番号（携帯、直通、Mobile、個人名の近くにある番号）
- fax: FAX番号

電話番号について：
- 複数の電話番号がある場合は、会社の代表番号と個人の番号を区別してください
- 「TEL」「代表」などの表記がある場合はcompanyPhoneに、「携帯」「Mobile」「直通」などの表記がある場合はpersonalPhoneに分類してください
- 電話番号が1つしかない場合は、文脈から判断して適切な方に分類してください（会社名の近くならcompanyPhone、個人名の近くならpersonalPhone）
- 判断が難しい場合は、companyPhoneとして扱ってください

フリガナは必ずカタカナ（例：タナカ タロウ、カブシキガイシャ サンプル）で出力してください。
存在しない項目は出力しないでください。JSONのみを返してください。`;

      const result = await this.model.generateContent([prompt, image]);
      const response = result.response;
      const text = response.text();
      
      const cleanText = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const parsedData = JSON.parse(cleanText);
      
      // 会社名フリガナから法人格を除去
      if (parsedData.companyFurigana && parsedData.companyName) {
        parsedData.companyFurigana = this.removeCompanyLegalForm(parsedData.companyName, parsedData.companyFurigana);
      }
      
      return {
        fileId,
        ...parsedData
      };
    } catch (error) {
      console.error('AI解析エラー:', error);
      return { fileId };
    }
  }

  async saveToSheet(spreadsheetId: string, data: BusinessCardData): Promise<void> {
    try {
      const values = [
        data.fileId,
        data.name || '',
        data.furigana || '',
        data.companyName || '',
        data.companyFurigana || '',
        data.department || '',
        data.position || '',
        data.postalCode || '',
        data.address || '',
        data.email || '',
        data.companyPhone || '',
        data.personalPhone || '',
        data.fax || ''
      ];

      await this.sheets.spreadsheets.values.append({
        spreadsheetId,
        range: `'${this.SHEET_NAME}'!A:M`,
        valueInputOption: 'RAW',
        resource: {
          values: [values]
        }
      });

      console.log(`データを保存しました: ${data.fileId}`);
    } catch (error) {
      console.error('データ保存エラー:', error);
      throw error;
    }
  }

  async moveFileToProcessed(fileId: string): Promise<void> {
    try {
      await this.drive.files.update({
        fileId,
        addParents: this.PROCESSED_FOLDER_ID,
        removeParents: this.BUSINESS_CARD_FOLDER_ID
      });

      console.log(`ファイルを処理済みフォルダに移動しました: ${fileId}`);
    } catch (error) {
      console.error('ファイル移動エラー:', error);
      throw error;
    }
  }

  async processBusinessCards(spreadsheetId: string): Promise<void> {
    try {
      console.log('名刺処理を開始します...');
      
      const allFileIds = await this.getFilesInFolder(this.BUSINESS_CARD_FOLDER_ID);
      const existingFileIds = await this.getExistingFileIds(spreadsheetId);
      const newFileIds = allFileIds.filter(fileId => !existingFileIds.includes(fileId));
      
      console.log(`フォルダ内ファイル数: ${allFileIds.length}`);
      console.log(`新規処理対象ファイル数: ${newFileIds.length}`);
      console.log(`重複ファイル数: ${allFileIds.length - newFileIds.length}`);

      if (allFileIds.length === 0) {
        console.log('フォルダ内にファイルがありません。');
        return;
      }

      // 新規ファイルの処理
      for (const fileId of newFileIds) {
        try {
          console.log(`処理中: ${fileId}`);
          
          const businessCardData = await this.analyzeBusinessCard(fileId);
          
          // 名前と会社名による重複チェック
          const isDuplicate = await this.isDuplicateData(spreadsheetId, businessCardData);
          
          if (isDuplicate) {
            console.log(`重複データのためスキップ: ${fileId} (名前: ${businessCardData.name}, 会社: ${businessCardData.companyName})`);
          } else {
            await this.saveToSheet(spreadsheetId, businessCardData);
            console.log(`データを保存: ${fileId}`);
          }
          
          await this.moveFileToProcessed(fileId);
          console.log(`完了: ${fileId}`);
        } catch (error) {
          console.error(`ファイル処理エラー ${fileId}:`, error);
        }
      }

      // 重複ファイルの移動
      const duplicateFileIds = allFileIds.filter(fileId => existingFileIds.includes(fileId));
      for (const fileId of duplicateFileIds) {
        try {
          console.log(`重複ファイルを移動中: ${fileId}`);
          await this.moveFileToProcessed(fileId);
          console.log(`重複ファイル移動完了: ${fileId}`);
        } catch (error) {
          console.error(`重複ファイル移動エラー ${fileId}:`, error);
        }
      }

      console.log('名刺処理が完了しました。');
    } catch (error) {
      console.error('処理エラー:', error);
      throw error;
    }
  }

  async startScheduledProcessing(spreadsheetId: string): Promise<void> {
    console.log('30分間隔での自動処理を開始します...');
    
    const processInterval = async () => {
      await this.processBusinessCards(spreadsheetId);
    };

    await processInterval();
    setInterval(processInterval, 30 * 60 * 1000);
  }
}

async function main() {
  const spreadsheetId = process.argv[2] || '1_aS8cRFajNlrAnvSgjmNhhM8U8W8lmvbbBDxmbJ_w6Q';

  const processor = new BusinessCardProcessor();
  
  try {
    await processor.initializeSheet(spreadsheetId);
    await processor.startScheduledProcessing(spreadsheetId);
  } catch (error) {
    console.error('実行エラー:', error);
    process.exit(1);
  }
}

if (require.main === module) {
  main().catch(console.error);
}

export { BusinessCardProcessor };