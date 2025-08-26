"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || (function () {
    var ownKeys = function(o) {
        ownKeys = Object.getOwnPropertyNames || function (o) {
            var ar = [];
            for (var k in o) if (Object.prototype.hasOwnProperty.call(o, k)) ar[ar.length] = k;
            return ar;
        };
        return ownKeys(o);
    };
    return function (mod) {
        if (mod && mod.__esModule) return mod;
        var result = {};
        if (mod != null) for (var k = ownKeys(mod), i = 0; i < k.length; i++) if (k[i] !== "default") __createBinding(result, mod, k[i]);
        __setModuleDefault(result, mod);
        return result;
    };
})();
Object.defineProperty(exports, "__esModule", { value: true });
exports.BusinessCardProcessor = void 0;
const generative_ai_1 = require("@google/generative-ai");
const googleapis_1 = require("googleapis");
const dotenv = __importStar(require("dotenv"));
dotenv.config();
class BusinessCardProcessor {
    constructor() {
        this.BUSINESS_CARD_FOLDER_ID = '1Q29bAIoQ__PADA2NefymTpsOO_yAg9ee';
        this.PROCESSED_FOLDER_ID = '16LAj4yAM2cyk-tUlY7WYFTkbmB5Krvjx';
        this.SHEET_NAME = '名刺情報';
        this.ai = new generative_ai_1.GoogleGenerativeAI(process.env.GEMINI_API_KEY || '');
        this.model = this.ai.getGenerativeModel({ model: 'gemini-2.5-pro' });
        const auth = new googleapis_1.google.auth.GoogleAuth({
            keyFile: process.env.GOOGLE_SERVICE_ACCOUNT_KEY_FILE,
            scopes: [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        });
        this.sheets = googleapis_1.google.sheets({ version: 'v4', auth });
        this.drive = googleapis_1.google.drive({ version: 'v3', auth });
    }
    async initializeSheet(spreadsheetId) {
        try {
            // まずシートの存在確認
            const sheetInfo = await this.sheets.spreadsheets.get({
                spreadsheetId,
                fields: 'sheets.properties'
            });
            const sheetExists = sheetInfo.data.sheets?.some((sheet) => sheet.properties?.title === this.SHEET_NAME);
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
                '部署', '肩書', '郵便番号', '住所', 'メールアドレス'
            ];
            await this.sheets.spreadsheets.values.update({
                spreadsheetId,
                range: `'${this.SHEET_NAME}'!A1:J1`,
                valueInputOption: 'RAW',
                resource: {
                    values: [headers]
                }
            });
            console.log('スプレッドシートのヘッダーを作成しました');
        }
        catch (error) {
            console.error('ヘッダー作成エラー:', error);
            throw error;
        }
    }
    async getFilesInFolder(folderId) {
        try {
            const response = await this.drive.files.list({
                q: `'${folderId}' in parents and trashed=false and (mimeType contains 'image/')`,
                fields: 'files(id, name)'
            });
            return response.data.files?.map((file) => file.id) || [];
        }
        catch (error) {
            console.error('フォルダ内ファイル取得エラー:', error);
            throw error;
        }
    }
    async getExistingFileIds(spreadsheetId) {
        try {
            const response = await this.sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `'${this.SHEET_NAME}'!A:A`
            });
            const values = response.data.values || [];
            return values.slice(1).map((row) => row[0]).filter(Boolean);
        }
        catch (error) {
            console.error('既存ファイルID取得エラー:', error);
            return [];
        }
    }
    async getExistingData(spreadsheetId) {
        try {
            const response = await this.sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `'${this.SHEET_NAME}'!B:D`
            });
            const values = response.data.values || [];
            return values.slice(1).map((row) => ({
                name: (row[0] || '').trim(),
                companyName: (row[2] || '').trim()
            }));
        }
        catch (error) {
            console.error('既存データ取得エラー:', error);
            return [];
        }
    }
    async isDuplicateData(spreadsheetId, data) {
        try {
            const existingData = await this.getExistingData(spreadsheetId);
            const newName = (data.name || '').trim();
            const newCompanyName = (data.companyName || '').trim();
            // 名前が空の場合は重複チェックをスキップ
            if (!newName) {
                return false;
            }
            // 名前の一致をチェック
            const nameMatch = existingData.find(existing => existing.name === newName);
            if (!nameMatch) {
                return false;
            }
            // 名前が一致した場合のみ会社名をチェック
            return nameMatch.companyName === newCompanyName;
        }
        catch (error) {
            console.error('重複チェックエラー:', error);
            return false;
        }
    }
    async getNewFiles(spreadsheetId) {
        const folderFileIds = await this.getFilesInFolder(this.BUSINESS_CARD_FOLDER_ID);
        const existingFileIds = await this.getExistingFileIds(spreadsheetId);
        return folderFileIds.filter(fileId => !existingFileIds.includes(fileId));
    }
    async downloadImageAsBase64(fileId) {
        try {
            const response = await this.drive.files.get({
                fileId,
                alt: 'media'
            }, { responseType: 'arraybuffer' });
            return Buffer.from(response.data).toString('base64');
        }
        catch (error) {
            console.error('画像ダウンロードエラー:', error);
            throw error;
        }
    }
    async analyzeBusinessCard(fileId) {
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

フリガナは必ずカタカナ（例：タナカ タロウ、カブシキガイシャ サンプル）で出力してください。
存在しない項目は出力しないでください。JSONのみを返してください。`;
            const result = await this.model.generateContent([prompt, image]);
            const response = result.response;
            const text = response.text();
            const cleanText = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
            const parsedData = JSON.parse(cleanText);
            return {
                fileId,
                ...parsedData
            };
        }
        catch (error) {
            console.error('AI解析エラー:', error);
            return { fileId };
        }
    }
    async saveToSheet(spreadsheetId, data) {
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
                data.email || ''
            ];
            await this.sheets.spreadsheets.values.append({
                spreadsheetId,
                range: `'${this.SHEET_NAME}'!A:J`,
                valueInputOption: 'RAW',
                resource: {
                    values: [values]
                }
            });
            console.log(`データを保存しました: ${data.fileId}`);
        }
        catch (error) {
            console.error('データ保存エラー:', error);
            throw error;
        }
    }
    async moveFileToProcessed(fileId) {
        try {
            await this.drive.files.update({
                fileId,
                addParents: this.PROCESSED_FOLDER_ID,
                removeParents: this.BUSINESS_CARD_FOLDER_ID
            });
            console.log(`ファイルを処理済みフォルダに移動しました: ${fileId}`);
        }
        catch (error) {
            console.error('ファイル移動エラー:', error);
            throw error;
        }
    }
    async processBusinessCards(spreadsheetId) {
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
                    }
                    else {
                        await this.saveToSheet(spreadsheetId, businessCardData);
                        console.log(`データを保存: ${fileId}`);
                    }
                    await this.moveFileToProcessed(fileId);
                    console.log(`完了: ${fileId}`);
                }
                catch (error) {
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
                }
                catch (error) {
                    console.error(`重複ファイル移動エラー ${fileId}:`, error);
                }
            }
            console.log('名刺処理が完了しました。');
        }
        catch (error) {
            console.error('処理エラー:', error);
            throw error;
        }
    }
    async startScheduledProcessing(spreadsheetId) {
        console.log('30分間隔での自動処理を開始します...');
        const processInterval = async () => {
            await this.processBusinessCards(spreadsheetId);
        };
        await processInterval();
        setInterval(processInterval, 30 * 60 * 1000);
    }
}
exports.BusinessCardProcessor = BusinessCardProcessor;
async function main() {
    const spreadsheetId = process.argv[2] || '1_aS8cRFajNlrAnvSgjmNhhM8U8W8lmvbbBDxmbJ_w6Q';
    const processor = new BusinessCardProcessor();
    try {
        await processor.initializeSheet(spreadsheetId);
        await processor.startScheduledProcessing(spreadsheetId);
    }
    catch (error) {
        console.error('実行エラー:', error);
        process.exit(1);
    }
}
if (require.main === module) {
    main().catch(console.error);
}
