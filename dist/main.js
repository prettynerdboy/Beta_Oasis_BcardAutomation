"use strict";
// To run this code you need to install the following dependencies:
// npm install @google/generative-ai dotenv
// npm install -D @types/node
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
const generative_ai_1 = require("@google/generative-ai");
const dotenv = __importStar(require("dotenv"));
const fs = __importStar(require("fs"));
const path = __importStar(require("path"));
dotenv.config();
async function main() {
    const ai = new generative_ai_1.GoogleGenerativeAI(process.env.GEMINI_API_KEY || '');
    const model = ai.getGenerativeModel({
        model: 'gemini-2.5-pro'
    });
    // コマンドライン引数から画像パスを取得
    const imagePath = process.argv[2];
    if (!imagePath) {
        console.log('使用方法: node dist/main.js <画像ファイルパス>');
        console.log('例: node dist/main.js ./BcardSample/sample1.jpg');
        process.exit(1);
    }
    // ファイルの存在確認
    if (!fs.existsSync(imagePath)) {
        console.error(`エラー: ファイルが見つかりません: ${imagePath}`);
        process.exit(1);
    }
    // 画像ファイルを読み込む
    const imageData = fs.readFileSync(imagePath);
    const mimeType = getMimeType(imagePath);
    const image = {
        inlineData: {
            data: imageData.toString('base64'),
            mimeType: mimeType
        }
    };
    const prompt = `名刺画像から以下の情報を抽出してJSONで返してください：
  - imageName: 画像データ名
  - name: 名前
  - furigana: 名前のフリガナ（カタカナで出力してください）
  - companyName: 社名
  - companyFurigana: 社名のフリガナ（カタカナで出力してください）
  - department: 部署
  - position: 肩書
  - postalCode: 郵便番号
  - address: 住所
  
  フリガナは必ずカタカナ（例：タナカ タロウ、カブシキガイシャ サンプル）で出力してください。
  存在しない項目は出力しないでください。JSONのみを返してください。`;
    const result = await model.generateContent([prompt, image]);
    const response = result.response;
    const text = response.text();
    // 画像ファイル名を取得
    const fileName = path.basename(imagePath);
    // JSON形式で整形して出力
    try {
        // ```json や ``` を削除してJSONだけを抽出
        const cleanText = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
        const json = JSON.parse(cleanText);
        // 画像データ名を追加
        json.imageName = fileName;
        console.log(JSON.stringify(json, null, 2));
    }
    catch {
        console.log(text);
    }
}
// MIMEタイプを判定する関数
function getMimeType(filePath) {
    const ext = path.extname(filePath).toLowerCase();
    const mimeTypes = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.webp': 'image/webp'
    };
    return mimeTypes[ext] || 'image/jpeg';
}
main().catch(console.error);
