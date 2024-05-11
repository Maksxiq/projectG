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
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
// Подключаем и импортируем необходимые библиотеки
const google_spreadsheet_1 = require("google-spreadsheet");
const googleapis_1 = require("googleapis");
const fs = __importStar(require("fs"));
const mammoth = __importStar(require("mammoth"));
// Пример использования функций
const sheetId = "your_google_sheet_id";
const docId = "your_google_doc_id";
const credentials = require('credentials.json');
// Функция для подключения к Google Sheets
function connectToGoogleSheet(sheetId, credentials) {
    return __awaiter(this, void 0, void 0, function* () {
        const doc = new google_spreadsheet_1.GoogleSpreadsheet(sheetId);
        try {
            yield doc.useServiceAccountAuth(credentials);
            yield doc.loadInfo(); // Загрузка информации о документе
            return doc;
        }
        catch (error) {
            console.error("Ошибка при подключении к Google Sheets:", error);
            return null;
        }
    });
}
function findSheetID(sheetLink) {
    const startIndex = sheetLink.indexOf("https://docs.google.com/spreadsheets/d/") + "https://docs.google.com/spreadsheets/d/".length;
    const endIndex = sheetLink.indexOf("/edit#gid=0", startIndex);
    if (startIndex !== -1 && endIndex !== -1) {
        return sheetLink.substring(startIndex, endIndex);
    }
    else {
        return null;
    }
}
// Функция для подключения к Google Docs
function connectToGoogleDocs(docId, credentials) {
    return __awaiter(this, void 0, void 0, function* () {
        const auth = new googleapis_1.google.auth.GoogleAuth({
            credentials: credentials,
            scopes: ['https://www.googleapis.com/auth/documents.readonly'],
        });
        const docs = googleapis_1.google.docs({ version: 'v1', auth });
        try {
            const res = yield docs.documents.get({
                documentId: docId,
            });
            return res.data;
        }
        catch (error) {
            console.error("Ошибка при подключении к Google Docs:", error);
            return null;
        }
    });
}
function findDocID(docLink) {
    const startIndex = docLink.indexOf("https://docs.google.com/document/d/") + "https://docs.google.com/document/d/".length;
    const endIndex = docLink.indexOf("/edit", startIndex);
    if (startIndex !== -1 && endIndex !== -1) {
        return docLink.substring(startIndex, endIndex);
    }
    else {
        return null;
    }
}
// Подключение к Google Sheets
const sheet = await connectToGoogleSheet(sheetId, credentials);
if (sheet) {
    console.log("Успешное подключение к Google Sheets:", sheet.title);
}
else {
    console.log("Не удалось подключиться к Google Sheets.");
}
// Подключение к Google Docs
const doc = await connectToGoogleDocs(docId, credentials);
if (doc) {
    console.log("Успешное подключение к Google Docs:", doc.title);
}
else {
    console.log("Не удалось подключиться к Google Docs.");
}
// Функция для извлечения значений переменных из документа .docx
function extractVariableValuesFromDocx(docxFile) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const result = yield mammoth.extractRawText({ path: docxFile });
            const text = result.value;
            const variableValues = new Map();
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
        }
        catch (error) {
            console.error("Ошибка при извлечении значений переменных из .docx:", error);
            return new Map();
        }
    });
}
// Функция для загрузки значений переменных из Google Sheets
function extractVariableValuesFromGoogleSheet(sheetId, credentials) {
    return __awaiter(this, void 0, void 0, function* () {
        const doc = new google_spreadsheet_1.GoogleSpreadsheet(sheetId);
        const variableValues = new Map();
        try {
            // Авторизация
            yield doc.useServiceAccountAuth(credentials);
            yield doc.loadInfo(); // Загрузка информации о документе
            const sheet = doc.sheetsByIndex[0]; // Предполагается, что данные находятся на первом листе
            // Получение данных из Google Sheets
            const rows = yield sheet.getRows();
            rows.forEach(row => {
                // Предполагаем, что первый столбец содержит названия переменных, а второй - их значения
                const variable = row._rawData[0].toString().trim();
                const value = row._rawData[1].toString().trim();
                variableValues.set(variable, value);
            });
            return variableValues;
        }
        catch (error) {
            console.error("Ошибка при извлечении значений переменных из Google Sheets:", error);
            return variableValues;
        }
    });
}
// Функция для создания нового документа Word с переменными
function createWordDocumentWithVariables(templatePath, outputPath, variableValues) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            const template = yield mammoth.convertToHtml({ path: templatePath });
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
        }
        catch (error) {
            console.error("Ошибка при создании документа Word:", error);
        }
    });
}
// Функция для скачивания файла по HTTP
function downloadFile(url, dest) {
    return new Promise((resolve, reject) => {
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
