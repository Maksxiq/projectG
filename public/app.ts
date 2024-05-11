 // Подключаем и импортируем необходимые библиотеки
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { google } from 'googleapis';
import { Credentials } from 'google-auth-library';
import * as fs from 'fs';
import * as mammoth from "mammoth";

// Пример использования функций
const sheetId = "your_google_sheet_id";
const docId = "your_google_doc_id";
const credentials = require('credentials.json');

// Функция для подключения к Google Sheets
async function connectToGoogleSheet(sheetId: string, credentials: Credentials) {
    const doc = new GoogleSpreadsheet(sheetId);
    try {
        await doc.useServiceAccountAuth(credentials);
        await doc.loadInfo(); // Загрузка информации о документе
        return doc;
    } catch (error) {
        console.error("Ошибка при подключении к Google Sheets:", error);
        return null;
    }
}

function findSheetID(sheetLink: string): string | null {
  const startIndex = sheetLink.indexOf("https://docs.google.com/spreadsheets/d/") + "https://docs.google.com/spreadsheets/d/".length;
  const endIndex = sheetLink.indexOf("/edit#gid=0", startIndex);
  if (startIndex !== -1 && endIndex !== -1) {
      return sheetLink.substring(startIndex, endIndex);
  } else {
      return null;
  }
}

// Функция для подключения к Google Docs
async function connectToGoogleDocs(docId: string, credentials: Credentials) {
    const auth = new google.auth.GoogleAuth({
        credentials: credentials,
        scopes: ['https://www.googleapis.com/auth/documents.readonly'],
    });
    const docs = google.docs({ version: 'v1', auth });
    try {
        const res = await docs.documents.get({
            documentId: docId,
        });
        return res.data;
    } catch (error) {
        console.error("Ошибка при подключении к Google Docs:", error);
        return null;
    }
}

function findDocID(docLink: string): string | null {
  const startIndex = docLink.indexOf("https://docs.google.com/document/d/") + "https://docs.google.com/document/d/".length;
  const endIndex = docLink.indexOf("/edit", startIndex);
  if (startIndex !== -1 && endIndex !== -1) {
      return docLink.substring(startIndex, endIndex);
  } else {
      return null;
  }
}

// Подключение к Google Sheets
const sheet = await connectToGoogleSheet(sheetId, credentials);
if (sheet) {
    console.log("Успешное подключение к Google Sheets:", sheet.title);
} else {
    console.log("Не удалось подключиться к Google Sheets.");
}

// Подключение к Google Docs
const doc = await connectToGoogleDocs(docId, credentials);
if (doc) {
    console.log("Успешное подключение к Google Docs:", doc.title);
} else {
    console.log("Не удалось подключиться к Google Docs.");
}

// Функция для извлечения значений переменных из документа .docx
async function extractVariableValuesFromDocx(docxFile: string): Promise<Map<string, string>> {
  try {
      const result = await mammoth.extractRawText({ path: docxFile });
      const text = result.value;
      const variableValues = new Map<string, string>();

      // Разделите текст на абзацы и найдите значения переменных
      const paragraphs = text.split("\n");
      paragraphs.forEach(paragraph => {
          const parts = paragraph.split(":"); // Предполагаем, что переменные имеют формат "название_переменной: значение_переменной"
          if (parts.length === 2) {
              const variable = parts[0].trim();
              const value = parts[1].trim();
              variableValues.set(variable, value);
          }
      });

      return variableValues;
  } catch (error) {
      console.error("Ошибка при извлечении значений переменных из .docx:", error);
      return new Map<string, string>();
  }
}

// Функция для загрузки значений переменных из Google Sheets
async function extractVariableValuesFromGoogleSheet(sheetId: string, credentials: Credentials): Promise<Map<string, string>> {
    const doc = new GoogleSpreadsheet(sheetId);
    const variableValues = new Map<string, string>();

    try {
        // Авторизация
        await doc.useServiceAccountAuth(credentials);
        await doc.loadInfo(); // Загрузка информации о документе
        const sheet = doc.sheetsByIndex[0]; // Предполагается, что данные находятся на первом листе

        // Получение данных из Google Sheets
        const rows = await sheet.getRows();
        rows.forEach(row => {
            // Предполагаем, что первый столбец содержит названия переменных, а второй - их значения
            const variable = row._rawData[0].toString().trim();
            const value = row._rawData[1].toString().trim();
            variableValues.set(variable, value);
        });

        return variableValues;
    } catch (error) {
        console.error("Ошибка при извлечении значений переменных из Google Sheets:", error);
        return variableValues;
    }
}

// Функция для создания нового документа Word с переменными
async function createWordDocumentWithVariables(templatePath: string, outputPath: string, variableValues: Map<string, string>): Promise<void> {
  try {
      const template = await mammoth.convertToHtml({ path: templatePath });
      let html = template.value;

      // Заменяем переменные в HTML документе значениями из Google Sheets
      variableValues.forEach((value, variable) => {
          const regex = new RegExp(`{{${variable}}}`, 'g');
          html = html.replace(regex, value);
      });

      // Сохраняем HTML в файл
      const htmlFilePath = path.join(__dirname, 'temp.html');
      fs.writeFileSync(htmlFilePath, html, 'utf8');

      // Конвертируем HTML в документ Word
      exec(`pandoc ${htmlFilePath} -o ${outputPath}`, (error, stdout, stderr) => {
          if (error) {
              console.error(`Ошибка при создании документа Word: ${error.message}`);
              return;
          }
          if (stderr) {
              console.error(`Ошибка при создании документа Word: ${stderr}`);
              return;
          }
          console.log(`Документ Word успешно создан: ${outputPath}`);
          // Открываем документ Word в модальном окне
          exec(`start ${outputPath}`);
      });
  } catch (error) {
      console.error("Ошибка при создании документа Word:", error);
  }
}

// Функция для скачивания файла по HTTP
function downloadFile(url: string, dest: string): Promise<void> {
  return new Promise<void>((resolve, reject) => {
      const file = fs.createWriteStream(dest);
      http.get(url, response => {
          response.pipe(file);
          file.on('finish', () => {
              file.close();
              resolve();
          });
      }).on('error', error => {
          fs.unlink(dest, () => {
              reject(error);
          });
      });
  });
}