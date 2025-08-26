// To run this code you need to install the following dependencies:
// npm install @google/generative-ai dotenv
// npm install -D @types/node

import {
  GoogleGenerativeAI,
  SchemaType,
} from '@google/generative-ai';
import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';

dotenv.config();

async function main() {
  const ai = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || '');
  
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

  } catch {
    console.log(text);
  }
}

// MIMEタイプを判定する関数
function getMimeType(filePath: string): string {
  const ext = path.extname(filePath).toLowerCase();
  const mimeTypes: { [key: string]: string } = {
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.png': 'image/png',
    '.gif': 'image/gif',
    '.webp': 'image/webp'
  };
  return mimeTypes[ext] || 'image/jpeg';
}

main().catch(console.error);