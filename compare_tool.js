#!/usr/bin/env node
/**
 * Инструмент сравнения таблицы уставок и файла конфигурации.
 *
 * Режим БЭ2704 (Nari):
 *   compare_tool.exe <таблица.docx> <конфиг.xml> [-o отчёт.txt] [--group N]
 *
 * Режим Siemens:
 *   compare_tool.exe <таблица.xlsx> <конфиг.xml> [-o отчёт.txt]
 *   (SiemensPie запускается автоматически; .xrio-файл ищется рядом с .xml по тому же имени)
 *
 *   compare_tool.exe          (интерактивный режим — запросит пути файлов)
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const readline = require('readline');
const { spawnSync } = require('child_process');

const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');

// Папка SiemensPie: в pkg-режиме — рядом с .exe, в dev-режиме — рядом со скриптом
const SIEMENSPIE_DIR = typeof process.pkg !== 'undefined'
    ? path.join(path.dirname(process.execPath), 'SiemensPie')
    : path.join(__dirname, 'SiemensPie');

// ─────────────────────────────────────────────────────────────────────────────
// Утилиты
// ─────────────────────────────────────────────────────────────────────────────

const ID_RE = /\[(\d{6})\]/;
/** Адрес параметра Siemens: 4 hex-цифры (+ опциональный буквенный суффикс) или 0x... */
const SIEMENS_DADR_RE = /^([0-9A-Fa-f]{4}[A-Za-z]?|0x[0-9A-Fa-f]+)$/;

/** Извлечь минимальный индекс из описания диапазона вида "(1 - выведено; 2 - введено)" */
function parseRangeMin(nameText) {
    const m = nameText.match(/\(([^)]+)\)/);
    if (!m) return 0;
    const inner = m[1];
    const nums = [...inner.matchAll(/(-?\d+)\s*[-–—]/g)].map(x => parseInt(x[1]));
    return nums.length ? Math.min(...nums) : 0;
}

/** Нормализация ё → е для устойчивого сравнения */
function normalizeEyo(s) {
    return s.replace(/ё/g, 'е');
}

/** Нормализация значения для сравнения */
function normalizeValue(val) {
    if (val === null || val === undefined) return '';
    let s = String(val).trim().replace(/[\n\r\t]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (!s) return '';
    // "-", "–", "—", "не используется" и строки-заметки (заканчиваются на ":") = пустое
    if (/^[-–—]+$/.test(s) || /^не\s+используется$/i.test(s) || /:\s*$/.test(s)) return '';
    // "N - описание" → берём только N
    const enumM = s.match(/^(-?\d+(?:[.,]\d+)?)\s*[-–—]\s*\S/);
    if (enumM) s = enumM[1];
    // Числовое сравнение
    const f = parseFloat(s.replace(',', '.').replace(/\s/g, ''));
    if (!isNaN(f)) {
        return Number.isInteger(f) ? String(Math.trunc(f)) : String(parseFloat(f.toPrecision(10)));
    }
    return s.toLowerCase().trim();
}

// ─────────────────────────────────────────────────────────────────────────────
// Парсинг XML-конфига
// ─────────────────────────────────────────────────────────────────────────────

function parseXmlConfig(xmlPath, group = 1) {
    const xmlText = fs.readFileSync(xmlPath, 'utf8');
    const parser  = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '@_',
        parseTagValue: false,
        parseAttributeValue: false,
    });
    const doc     = parser.parse(xmlText);

    // Корневой элемент может называться по-разному
    const rootKey = Object.keys(doc).find(k => k !== '?xml');
    const root    = doc[rootKey] || doc;
    const params  = {};

    // ── Ratios ────────────────────────────────────────────────────────────────
    const ratiosNode = root.Ratios;
    if (ratiosNode) {
        const items = Array.isArray(ratiosNode.Ratio) ? ratiosNode.Ratio : [ratiosNode.Ratio];
        for (const r of items) {
            if (!r) continue;
            const fullId = r['@_ID'] || '';
            const m = fullId.match(/\d+:(\d+)/);
            if (!m) continue;
            const nid = m[1];
            const val = r['@_Value'] || '';
            params[nid] = {
                id: nid, fullId, name: `[${nid}]`,
                value: val, unit: r['@_Unit'] || '',
                source: 'Ratios', path: 'Коэффициенты трансформации',
                allValues: { 1: val },
            };
        }
    }

    // ── Initials ──────────────────────────────────────────────────────────────
    const initialsNode = root.Initials;
    if (initialsNode) {
        const items = Array.isArray(initialsNode.Initial) ? initialsNode.Initial : [initialsNode.Initial];
        for (const r of items) {
            if (!r) continue;
            const fullId = r['@_ID'] || '';
            const m = fullId.match(/\d+:(\d+)/);
            if (!m) continue;
            const nid = m[1];
            const val = r['@_Value'] || '';
            params[nid] = {
                id: nid, fullId, name: `[${nid}]`,
                value: val, unit: '',
                source: 'Initials', path: 'Общие сведения',
                allValues: { 1: val },
            };
        }
    }

    // ── Settings (рекурсивно) ─────────────────────────────────────────────────
    function collectSettings(node, curPath) {
        if (!node) return;
        const nodeName = node['@_Name'] || '';
        const nextPath = curPath ? `${curPath} / ${nodeName}`.replace(/^\s*\/\s*/, '') : nodeName;

        // Дочерние Setting
        const settings = node.Setting
            ? (Array.isArray(node.Setting) ? node.Setting : [node.Setting])
            : [];
        for (const s of settings) {
            if (!s) continue;
            const fullId = s['@_ID'] || '';
            const m = fullId.match(/\d+:(\d+)/);
            if (!m) continue;
            const nid = m[1];
            const allVals = {};
            for (let g = 1; g <= 4; g++) {
                const v = s[`@_Value${g}`];
                if (v !== undefined) allVals[g] = v;
            }
            const value = s[`@_Value${group}`] !== undefined
                ? s[`@_Value${group}`]
                : (s['@_Value1'] !== undefined ? s['@_Value1'] : '');
            // Inline список возможных значений <PossibleValues><Item Value="..."/></PossibleValues>
            let possibleValues = [];
            if (s.PossibleValues) {
                const pvItems = s.PossibleValues.Item
                    ? (Array.isArray(s.PossibleValues.Item) ? s.PossibleValues.Item : [s.PossibleValues.Item])
                    : [];
                possibleValues = pvItems.map(it => String(it['@_Value'] || ''));
            }
            params[nid] = {
                id: nid, fullId,
                name: s['@_Name'] || `[${nid}]`,
                value: String(value),
                unit: s['@_Unit'] || '',
                ratio: s['@_Ratio'] || '',
                expander: s['@_PossibleValuesExpander'] || '',
                possibleValues,
                source: 'Settings',
                path: nextPath,
                allValues: allVals,
            };
        }

        // Дочерние Node
        const nodes = node.Node
            ? (Array.isArray(node.Node) ? node.Node : [node.Node])
            : [];
        for (const child of nodes) {
            collectSettings(child, nextPath);
        }
    }

    if (root.Settings) {
        const topNodes = root.Settings.Node
            ? (Array.isArray(root.Settings.Node) ? root.Settings.Node : [root.Settings.Node])
            : [];
        for (const n of topNodes) collectSettings(n, '');
    }

    // ── Expanders (именованные списки возможных значений) ─────────────────────
    const expanders = {};
    const expandersNode = root.Expanders;
    if (expandersNode) {
        const expList = Array.isArray(expandersNode.Expander)
            ? expandersNode.Expander
            : (expandersNode.Expander ? [expandersNode.Expander] : []);
        for (const exp of expList) {
            if (!exp) continue;
            const name = exp['@_Name'] || '';
            const items = exp.Item
                ? (Array.isArray(exp.Item) ? exp.Item : [exp.Item])
                : [];
            expanders[name] = items.map(it => String(it['@_Value'] || ''));
        }
    }

    return { params, expanders };
}

// ─────────────────────────────────────────────────────────────────────────────
// Парсинг Word (.docx)
// ─────────────────────────────────────────────────────────────────────────────

/** Извлечь текст из XML-элемента <w:t> рекурсивно */
function extractText(node) {
    if (!node) return '';
    if (typeof node === 'string') return node;
    if (typeof node === 'number') return String(node);
    let text = '';
    for (const key of Object.keys(node)) {
        if (key === '@_xml:space' || key.startsWith('@_')) continue;
        if (key === 'w:t' || key === 'w:delText') {
            const v = node[key];
            if (typeof v === 'string') text += v;
            else if (typeof v === 'object' && v['#text'] !== undefined) text += v['#text'];
            else if (typeof v === 'number') text += String(v);
        } else {
            const child = node[key];
            if (Array.isArray(child)) text += child.map(extractText).join('');
            else if (typeof child === 'object') text += extractText(child);
        }
    }
    return text;
}

function parseDocxTable(docxPath) {
    const zip = new AdmZip(docxPath);
    const xmlEntry = zip.getEntry('word/document.xml');
    if (!xmlEntry) throw new Error('Не найден word/document.xml в файле ' + docxPath);

    const xmlText = xmlEntry.getData().toString('utf8');
    const parser  = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '@_',
        parseTagValue: false,       // Не конвертировать текст в числа ("000" → не 0)
        parseAttributeValue: false,
        isArray: (tag) => ['w:tr', 'w:tc', 'w:tbl', 'w:r', 'w:p'].includes(tag),
    });
    const doc = parser.parse(xmlText);

    const body = doc?.['w:document']?.['w:body']
              || doc?.['w:wordDocument']?.['w:body']
              || doc?.['w:body'];
    if (!body) throw new Error('Не удалось найти <w:body> в документе');

    // Собрать все таблицы
    const tables = body['w:tbl']
        ? (Array.isArray(body['w:tbl']) ? body['w:tbl'] : [body['w:tbl']])
        : [];

    const params = {};

    for (const table of tables) {
        const rows = table['w:tr']
            ? (Array.isArray(table['w:tr']) ? table['w:tr'] : [table['w:tr']])
            : [];

        // Определить, есть ли в первых двух строках упоминание "первичные"
        let isPrimary = false;
        for (let ri = 0; ri < Math.min(2, rows.length); ri++) {
            const rowText = JSON.stringify(rows[ri]);
            if (/первичн/i.test(rowText)) { isPrimary = true; break; }
        }

        for (const row of rows) {
            const cells = row['w:tc']
                ? (Array.isArray(row['w:tc']) ? row['w:tc'] : [row['w:tc']])
                : [];

            // Извлечь текст каждой ячейки
            const cellTexts = cells.map(cell => {
                const paragraphs = cell['w:p']
                    ? (Array.isArray(cell['w:p']) ? cell['w:p'] : [cell['w:p']])
                    : [];
                return paragraphs.map(p => extractText(p)).join(' ')
                    .replace(/\s+/g, ' ').trim();
            });

            // Убрать дублирующиеся соседние ячейки (объединённые ячейки)
            const uniq = [];
            for (const t of cellTexts) {
                if (uniq.length === 0 || t !== uniq[uniq.length - 1]) uniq.push(t);
            }

            // Найти ячейку с [XXXXXX]
            let idIdx = -1;
            for (let i = 0; i < uniq.length; i++) {
                if (ID_RE.test(uniq[i])) { idIdx = i; break; }
            }
            if (idIdx === -1) continue;

            const nameText = uniq[idIdx];
            const m = nameText.match(ID_RE);
            if (!m) continue;
            const nid = m[1];

            const valueText = idIdx + 1 < uniq.length ? uniq[idIdx + 1] : '';
            const unitText  = idIdx + 2 < uniq.length ? uniq[idIdx + 2] : '';

            // Пропустить заголовки таблиц
            if (/значени|наименован|диапазон/i.test(valueText)) continue;

            if (!(nid in params)) {
                params[nid] = {
                    id: nid, name: nameText,
                    value: valueText, unit: unitText,
                    primaryValues: isPrimary,
                    rangeMin: parseRangeMin(nameText),
                };
            }
        }
    }

    return params;
}

// ─────────────────────────────────────────────────────────────────────────────
// Парсинг xlsx (таблица уставок Siemens / вывод SiemensPie)
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Прочитать таблицу параметров из xlsx-файла.
 * Ожидаемый формат (SiemensPie / таблица уставок Siemens):
 *   A = DAdr, B = короткое имя, C = диапазон/варианты, D = значение, H = комментарий
 * Строки без DAdr в колонке A — заголовки разделов.
 */
function parseXlsxParams(xlsxPath) {
    const zip = new AdmZip(xlsxPath);

    // ── Shared strings ────────────────────────────────────────────────────────
    const strings = [];
    const ssEntry = zip.getEntry('xl/sharedStrings.xml');
    if (ssEntry) {
        const ssText = ssEntry.getData().toString('utf8');
        const ssParser = new XMLParser({
            ignoreAttributes: false,
            attributeNamePrefix: '@_',
            parseTagValue: false,
            parseAttributeValue: false,
            isArray: tag => tag === 'si' || tag === 'r',
        });
        const ssDoc = ssParser.parse(ssText);
        const sst = ssDoc.sst || ssDoc;
        const siArr = Array.isArray(sst.si) ? sst.si : (sst.si ? [sst.si] : []);
        for (const si of siArr) {
            if (!si) { strings.push(''); continue; }
            if (si.t !== undefined) {
                const tv = si.t;
                strings.push(typeof tv === 'object' ? (tv['#text'] || '') : String(tv || ''));
            } else if (si.r) {
                const parts = Array.isArray(si.r) ? si.r : [si.r];
                strings.push(parts.map(r => {
                    if (!r || r.t === undefined) return '';
                    const tv = r.t;
                    return typeof tv === 'object' ? (tv['#text'] || '') : String(tv || '');
                }).join(''));
            } else {
                strings.push('');
            }
        }
    }

    // ── Выбор листа по имени из workbook.xml ─────────────────────────────────
    // Приоритет: "Config" (вывод SiemensPie) > "ТУ" (таблица уставок) > первый лист
    const sheetEntry = (() => {
        try {
            const wbEntry   = zip.getEntry('xl/workbook.xml');
            const relsEntry = zip.getEntry('xl/_rels/workbook.xml.rels');
            if (!wbEntry || !relsEntry) throw new Error('no workbook');

            const metaParser = new XMLParser({
                ignoreAttributes: false,
                attributeNamePrefix: '@_',
                parseTagValue: false,
                parseAttributeValue: false,
                isArray: tag => tag === 'sheet' || tag === 'Relationship',
            });
            const wbDoc   = metaParser.parse(wbEntry.getData().toString('utf8'));
            const relsDoc = metaParser.parse(relsEntry.getData().toString('utf8'));

            // rId → target path (относительно xl/)
            const ridToFile = {};
            const relsArr = relsDoc?.Relationships?.Relationship || [];
            for (const rel of (Array.isArray(relsArr) ? relsArr : [relsArr])) {
                const t = rel['@_Type'] || '';
                if (t.includes('worksheet')) {
                    ridToFile[rel['@_Id']] = 'xl/' + String(rel['@_Target']).replace(/^\//, '');
                }
            }

            // Упорядоченный список листов
            const wb = wbDoc?.workbook || wbDoc;
            const sheetsArr = wb?.sheets?.sheet || [];
            const sheets = Array.isArray(sheetsArr) ? sheetsArr : [sheetsArr];

            // Найти по имени (Config → ТУ → первый)
            const byName = {};
            let firstEntry = null;
            for (const sh of sheets) {
                const name = (sh['@_name'] || '').toLowerCase().trim();
                // r:id может прийти как @_r:id
                const rid  = sh['@_r:id'] || sh['@_r_id'] || '';
                const file = ridToFile[rid];
                if (!file) continue;
                const entry = zip.getEntry(file);
                if (!entry) continue;
                if (!firstEntry) firstEntry = entry;
                byName[name] = entry;
            }

            return byName['config'] || byName['ту'] || firstEntry;
        } catch (_) {
            // Fallback если workbook не удалось прочитать
            return zip.getEntry('xl/worksheets/sheet1.xml');
        }
    })();
    if (!sheetEntry) throw new Error(`Не найден лист с уставками в файле ${xlsxPath}`);

    const sheetText = sheetEntry.getData().toString('utf8');
    const shParser = new XMLParser({
        ignoreAttributes: false,
        attributeNamePrefix: '@_',
        parseTagValue: false,
        parseAttributeValue: false,
        isArray: tag => tag === 'row' || tag === 'c',
    });
    const shDoc = shParser.parse(sheetText);
    const ws = shDoc.worksheet || shDoc;
    const sd = ws.sheetData || {};
    const rows = Array.isArray(sd.row) ? sd.row : (sd.row ? [sd.row] : []);

    const params = {};
    let currentSection = '';

    for (const row of rows) {
        const cells = Array.isArray(row.c) ? row.c : (row.c ? [row.c] : []);
        const rd = {};

        for (const cell of cells) {
            const ref = String(cell['@_r'] || '');
            const col = ref.replace(/\d+$/, '');   // "A1" → "A", "AB12" → "AB"
            const type = String(cell['@_t'] || '');
            let val = '';

            if (type === 's') {
                const idx = parseInt(String(cell.v || ''));
                val = (!isNaN(idx) && idx < strings.length) ? strings[idx] : '';
            } else if (type === 'inlineStr') {
                val = cell.is && cell.is.t ? String(cell.is.t) : '';
            } else {
                val = String(cell.v || '');
            }
            // Убрать артефакты переноса строки из xlsx (_x000D_)
            val = val.replace(/_x000D_\r?\n?/g, '\n').trim();
            if (col) rd[col] = val;
        }

        const addrRaw = (rd['A'] || '').trim();

        // Строки без DAdr → заголовок раздела
        if (!SIEMENS_DADR_RE.test(addrRaw)) {
            if (addrRaw
                && addrRaw !== 'Адрес'
                && !/^\d+$/.test(addrRaw)
                && !addrRaw.startsWith('УTBEPЖДAЮ')
                && addrRaw.length > 3) {
                currentSection = addrRaw;
            }
            continue;
        }

        // Строка-заголовок колонок (col B = "Параметр")
        if ((rd['B'] || '').trim() === 'Параметр') continue;

        const shortName = (rd['B'] || '').trim();
        const comment   = (rd['H'] || '').trim();
        // Убрать завершающий '*' (признак примечания в таблице)
        const value = (rd['D'] || '').replace(/\*+$/, '').trim();

        params[addrRaw] = {
            id: addrRaw,
            name: `[${addrRaw}] ${shortName || comment || addrRaw}`.trim(),
            value,
            unit: '',
            path: currentSection,
            source: 'xlsx',
            // поля ниже нужны для совместимости с compare()
            primaryValues: false,
            rangeMin: 0,
        };
    }

    return params;
}

// ─────────────────────────────────────────────────────────────────────────────
// Сравнение
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Получить текстовую метку значения-селектора по индексу из конфига.
 * Пробует сначала inline PossibleValues, затем именованный Expander.
 */
function resolveXmlLabel(xp, expanders) {
    let vals = xp.possibleValues || [];
    if (vals.length === 0 && xp.expander && expanders[xp.expander]) {
        vals = expanders[xp.expander];
    }
    if (vals.length === 0) return null;
    const idx = parseInt(xp.value);
    if (isNaN(idx) || idx < 0 || idx >= vals.length) return null;
    return vals[idx];
}

/**
 * Вычислить относительное расхождение в процентах между двумя
 * нормализованными строковыми значениями. Возвращает null если
 * хотя бы одно значение не числовое.
 */
function calcPctDiff(wNorm, xNorm) {
    const w = parseFloat(wNorm);
    const x = parseFloat(xNorm);
    if (isNaN(w) || isNaN(x)) return null;
    const denom = Math.max(Math.abs(w), Math.abs(x));
    if (denom === 0) return 0;
    return Math.abs(w - x) / denom * 100;
}

function compare(wordParams, xmlParams, expanders) {
    const allIds = [...new Set([...Object.keys(wordParams), ...Object.keys(xmlParams)])].sort();
    const differences = [], smallDiff = [], onlyWord = [], onlyXml = [], tableEmpty = [];

    for (const nid of allIds) {
        const inWord = nid in wordParams;
        const inXml  = nid in xmlParams;

        if (inWord && !inXml) {
            onlyWord.push(wordParams[nid]);
        } else if (inXml && !inWord) {
            const xp = xmlParams[nid];
            if (xp.value && xp.value.trim()) onlyXml.push(xp);
        } else {
            const wp = wordParams[nid], xp = xmlParams[nid];
            const wn = normalizeValue(wp.value);

            // Правка 2: Пустое значение в таблице → отдельная категория
            if (!wn) {
                const xn = normalizeValue(xp.value);
                if (xn) {
                    tableEmpty.push({
                        id: nid, name: xp.name, path: xp.path || '',
                        valueXml: xp.value, unitXml: xp.unit || '',
                        allXmlValues: xp.allValues || {},
                    });
                }
                continue;
            }

            // Правка 3: Селекторы — вычесть смещение rangeMin для сравнения
            let wCompare = wn;
            const enumMatch = wp.value.trim().match(/^(-?\d+(?:[.,]\d+)?)\s*[-–—]/);
            const rangeMin = wp.rangeMin || 0;
            if (enumMatch && rangeMin !== 0) {
                const wNum = parseFloat(enumMatch[1].replace(',', '.'));
                if (!isNaN(wNum)) {
                    wCompare = normalizeValue(String(wNum - rangeMin));
                }
            }

            // Правка 1: Первичные/вторичные — перевод значения конфига в первичные
            let xValueForCompare = xp.value;
            let xValuePrimary = null;
            if (wp.primaryValues && xp.ratio && xp.ratio.trim()) {
                const ratioF = parseFloat(xp.ratio);
                const xNum = parseFloat(String(xp.value).replace(',', '.').replace(/\s/g, ''));
                if (!isNaN(ratioF) && !isNaN(xNum) && ratioF > 0) {
                    // toFixed(3) исключает ошибки округления на 10-тысячных
                    xValuePrimary = String(parseFloat((xNum * ratioF).toFixed(3)));
                    xValueForCompare = xValuePrimary;
                }
            }

            const xn = normalizeValue(xValueForCompare);
            if (!wn && !xn) continue;
            if (wCompare !== xn) {
                // Правка 3 (резервный путь): сопоставление через PossibleValues / Expander.
                // Работает когда таблица содержит текстовую метку ("вперёд"), а конфиг — индекс ("0").
                const xmlLabel = resolveXmlLabel(xp, expanders || {});
                if (xmlLabel !== null) {
                    const normLabel = normalizeEyo(normalizeValue(xmlLabel));
                    if (normalizeEyo(wn) === normLabel) continue; // совпадение по метке
                }
                const pct = calcPctDiff(wCompare, xn);
                const entry = {
                    id: nid, name: xp.name, path: xp.path || '',
                    valueWord: wp.value, unitWord: wp.unit || '',
                    valueXml: xp.value,  unitXml: xp.unit  || '',
                    valueXmlPrimary: xValuePrimary,
                    primaryValues: wp.primaryValues || false,
                    allXmlValues: xp.allValues || {},
                    pctDiff: pct,
                };
                // Числовое расхождение < 1% → отдельный раздел, не в основные различия
                if (pct !== null && pct < 1) {
                    smallDiff.push(entry);
                } else {
                    differences.push(entry);
                }
            }
        }
    }

    return { differences, smallDiff, onlyWord, onlyXml, tableEmpty };
}

// ─────────────────────────────────────────────────────────────────────────────
// Форматирование отчёта
// ─────────────────────────────────────────────────────────────────────────────

function formatReport(results, wordPath, xmlPath, wordParams, xmlParams, group, deviceType) {
    const SEP = '='.repeat(72);
    const lines = [];
    const add = (s = '') => lines.push(s);
    const isSiemens = deviceType === 'Siemens';

    const now = new Date().toLocaleString('ru-RU');
    add(SEP);
    add('ОТЧЁТ СРАВНЕНИЯ ТАБЛИЦЫ УСТАВОК И ФАЙЛА КОНФИГУРАЦИИ');
    add(SEP);
    add(`Сформирован:          ${now}`);
    if (isSiemens) {
        add(`Тип терминала:        Siemens`);
        add(`Таблица уставок:      ${wordPath}`);
        add(`Вывод SiemensPie:     ${xmlPath}`);
    } else {
        add(`Таблица уставок:      ${wordPath}`);
        add(`Файл конфигурации:    ${xmlPath}`);
        add(`Группа уставок XML:   ${group}`);
    }
    add();
    add(`Параметров в таблице:         ${Object.keys(wordParams).length}`);
    add(`Параметров в конфиге (всего): ${Object.keys(xmlParams).length}`);
    add();
    const { differences: diffs, smallDiff: sdiffs = [],
            onlyWord: ow, onlyXml: ox, tableEmpty: te = [] } = results;
    add(`Итог: Различий значительных: ${diffs.length}  |  Расхождение <1%: ${sdiffs.length}  |  Только в таблице: ${ow.length}  |  Только в конфиге: ${ox.length}  |  Пусто в таблице: ${te.length}`);
    add();

    /** Вывод одной строки различия (используется в разделах 1 и 2) */
    function addDiffEntry(d) {
        add(`  ${d.name}`);
        if (d.path) add(`    Раздел: ${d.path}`);
        const wStr = d.valueWord + (d.unitWord ? ` ${d.unitWord}` : '');
        const xRaw = d.valueXml  + (d.unitXml  ? ` ${d.unitXml}`  : '');
        add(`    Таблица:  ${wStr || '(пусто)'}`);
        if (d.primaryValues && d.valueXmlPrimary) {
            add(`    Конфиг:   ${xRaw || '(пусто)'}  →  ${d.valueXmlPrimary} ${d.unitXml} (первичные)`);
        } else {
            add(`    Конфиг:   ${xRaw || '(пусто)'}`);
            if (d.primaryValues) add(`    * Таблица: первичные, коэффициент трансформации не определён`);
        }
        if (d.pctDiff !== null && d.pctDiff !== undefined) {
            add(`    Расхождение: ${d.pctDiff.toFixed(2)}%`);
        }
        const av = d.allXmlValues;
        const avKeys = Object.keys(av);
        if (avKeys.length > 1) {
            add(`    Конфиг (все группы): ${avKeys.map(g => `Гр${g}=${av[g]}`).join('  ')}`);
        }
        add();
    }

    // 1. Значительные различия (≥1% или нечисловые)
    add(SEP);
    add(`1. РАЗЛИЧИЯ ЗНАЧЕНИЙ — ${diffs.length} позиций`);
    add(SEP);
    if (diffs.length) {
        for (const d of diffs) addDiffEntry(d);
    } else {
        add('  Различий нет.');
        add();
    }

    // 2. Малые расхождения (<1%)
    add(SEP);
    add(`2. РАСХОЖДЕНИЕ МЕНЕЕ 1% — ${sdiffs.length} позиций`);
    add(SEP);
    if (sdiffs.length) {
        for (const d of sdiffs) addDiffEntry(d);
    } else {
        add('  Нет таких параметров.');
        add();
    }

    const cfgLabel = isSiemens ? 'конфигурации (SiemensPie)' : 'конфиге';
    const cfgLabelShort = isSiemens ? 'конфигурации' : 'конфиге';

    // 3. Только в таблице
    add(SEP);
    add(`3. ЕСТЬ В ТАБЛИЦЕ, НЕТ В ${cfgLabel.toUpperCase()} — ${ow.length} позиций`);
    add(SEP);
    if (ow.length) {
        for (const p of ow) {
            add(`  ${p.name}`);
            const vStr = p.value + (p.unit ? ` ${p.unit}` : '');
            if (vStr.trim()) add(`    Значение в таблице: ${vStr}`);
            add();
        }
    } else {
        add('  Нет таких параметров.');
        add();
    }

    // 4. Только в конфиге
    add(SEP);
    add(`4. ЕСТЬ В ${cfgLabel.toUpperCase()} (со значением), НЕТ В ТАБЛИЦЕ — ${ox.length} позиций`);
    add(SEP);
    if (ox.length) {
        for (const p of ox) {
            add(`  ${p.name}`);
            if (p.path) add(`    Раздел: ${p.path}`);
            const vStr = p.value + (p.unit ? ` ${p.unit}` : '');
            add(`    Значение в ${cfgLabelShort}: ${vStr}`);
            add();
        }
    } else {
        add('  Нет таких параметров.');
        add();
    }

    // 5. Не заполнено в таблице
    add(SEP);
    add(`5. НЕ ЗАПОЛНЕНО В ТАБЛИЦЕ — ${te.length} позиций`);
    add(SEP);
    if (te.length) {
        for (const p of te) {
            add(`  ${p.name}`);
            if (p.path) add(`    Раздел: ${p.path}`);
            const vStr = p.valueXml + (p.unitXml ? ` ${p.unitXml}` : '');
            add(`    Значение в ${cfgLabelShort}: ${vStr}`);
            add();
        }
    } else {
        add('  Нет таких параметров.');
        add();
    }

    add(SEP);
    add('КОНЕЦ ОТЧЁТА');
    add(SEP);
    return lines.join('\r\n');
}

// ─────────────────────────────────────────────────────────────────────────────
// Главная функция
// ─────────────────────────────────────────────────────────────────────────────

/** Записать ошибку в лог-файл в текущей рабочей папке */
function writeErrorLog(err) {
    try {
        const logPath = path.join(process.cwd(), 'compare_tool_errors.log');
        const msg = `[${new Date().toISOString()}]\n${err && err.stack ? err.stack : err}\n\n`;
        fs.appendFileSync(logPath, msg, 'utf8');
        console.error(`Лог ошибок: ${logPath}`);
    } catch (_) { /* ignore */ }
}

/** Зависающий процесс — нажмите Enter */
async function pressEnterToExit() {
    const rl2 = readline.createInterface({ input: process.stdin, output: process.stdout });
    await new Promise(res => rl2.question('\nНажмите Enter для выхода...', res));
    rl2.close();
}

/** Убрать кавычки и лишние пробелы из пути введённого пользователем */
function cleanPath(p) {
    return p.trim().replace(/^["']|["']$/g, '').trim();
}

/** Общая часть: сохранить отчёт и открыть в Блокноте */
function saveAndOpen(report, basePath, outputPath) {
    const date = new Date().toISOString().slice(0, 10);
    const absOut = outputPath
        ? path.resolve(outputPath)
        : path.join(path.dirname(basePath), `report_${date}.txt`);
    fs.writeFileSync(absOut, '\uFEFF' + report, 'utf8');
    if (process.platform === 'win32') {
        try {
            require('child_process').spawn('notepad.exe', [absOut], { detached: true, stdio: 'ignore' }).unref();
        } catch (_) {}
    }
    return absOut;
}

function logSummary(results) {
    const d = results.differences.length;
    const s = (results.smallDiff || []).length;
    const w = results.onlyWord.length;
    const x = results.onlyXml.length;
    const e = results.tableEmpty.length;
    console.log(`Итог: Различий: ${d}  |  <1%: ${s}  |  Только в таблице: ${w}  |  Только в конфиге: ${x}  |  Пусто в таблице: ${e}`);
}

/** Режим БЭ2704: .docx vs .xml */
async function run(wordPath, xmlPath, outputPath, group) {
    const absWord = path.resolve(wordPath);
    const absXml  = path.resolve(xmlPath);

    if (!fs.existsSync(absWord)) throw new Error(`Файл не найден: ${absWord}`);
    if (!fs.existsSync(absXml))  throw new Error(`Файл не найден: ${absXml}`);

    console.log(`Таблица уставок:   ${absWord}`);
    console.log(`Файл конфигурации: ${absXml}`);
    console.log();

    console.log('Чтение таблицы уставок ...');
    const wordParams = parseDocxTable(absWord);
    console.log(`  Найдено параметров: ${Object.keys(wordParams).length}`);

    console.log('Чтение конфигурации ...');
    const { params: xmlParams, expanders } = parseXmlConfig(absXml, group);
    console.log(`  Найдено параметров: ${Object.keys(xmlParams).length}`);

    console.log('Сравнение ...');
    const results = compare(wordParams, xmlParams, expanders);
    const report  = formatReport(results, absWord, absXml, wordParams, xmlParams, group);
    const absOut  = saveAndOpen(report, absWord, outputPath);
    console.log(`\nОтчёт сохранён: ${absOut}`);
    logSummary(results);
    return absOut;
}

/**
 * Запустить SiemensPie (sp.exe) и вернуть путь к сгенерированному xlsx.
 * sp.exe + python37.dll + config.json должны лежать в папке SiemensPie/
 * рядом с compare_tool.exe (или рядом со скриптом в dev-режиме).
 */
function invokeSiemensPie(xmlPath, xrioPath) {
    const spExe = path.join(SIEMENSPIE_DIR, 'sp.exe');
    if (!fs.existsSync(spExe)) {
        throw new Error(
            `sp.exe не найден: ${spExe}\n` +
            `Убедитесь, что папка SiemensPie/ находится рядом с compare_tool.exe`
        );
    }
    const spArgs = xrioPath ? [xmlPath, xrioPath] : [xmlPath];
    console.log(`Запуск SiemensPie: ${path.basename(xmlPath)}${xrioPath ? ' + ' + path.basename(xrioPath) : ''} ...`);
    // cwd = папка SiemensPie, чтобы python37.dll и config.json были рядом с sp.exe
    const result = spawnSync(spExe, spArgs, { cwd: SIEMENSPIE_DIR, timeout: 60000 });
    if (result.stdout && result.stdout.length) process.stdout.write(result.stdout);
    if (result.stderr && result.stderr.length) process.stderr.write(result.stderr);
    if (result.error) throw result.error;
    if (result.status !== 0) throw new Error(`SiemensPie завершился с кодом ${result.status}`);
    // sp.exe создаёт xlsx рядом с xml-файлом, с тем же именем
    const outXlsx = path.join(
        path.dirname(xmlPath),
        path.basename(xmlPath, path.extname(xmlPath)) + '.xlsx'
    );
    if (!fs.existsSync(outXlsx)) throw new Error(`SiemensPie не создал файл: ${outXlsx}`);
    return outXlsx;
}

/** Режим Siemens: таблица уставок .xlsx + конфиг .xml → запускает SiemensPie, затем сравнивает */
async function runSiemens(tablePath, xmlPath, outputPath) {
    const absTable = path.resolve(tablePath);
    const absXml   = path.resolve(xmlPath);

    if (!fs.existsSync(absTable)) throw new Error(`Файл не найден: ${absTable}`);
    if (!fs.existsSync(absXml))   throw new Error(`Файл не найден: ${absXml}`);

    // Найти .xrio рядом с xml (то же имя, другое расширение)
    const xrioCandidate = absXml.replace(/\.xml$/i, '.xrio');
    const absXrio = fs.existsSync(xrioCandidate) ? xrioCandidate : null;
    if (!absXrio) console.warn(`  Предупреждение: .xrio-файл не найден рядом с ${path.basename(absXml)}, SiemensPie запустится без него.`);

    console.log(`Таблица уставок: ${absTable}`);
    console.log(`Конфиг XML:      ${absXml}`);
    if (absXrio) console.log(`Конфиг XRio:     ${absXrio}`);
    console.log();

    // Запустить SiemensPie → получить xlsx с уставками из конфига
    const absConfig = invokeSiemensPie(absXml, absXrio);
    console.log(`  Вывод SiemensPie: ${absConfig}`);
    console.log();

    console.log('Чтение таблицы уставок ...');
    const tableParams  = parseXlsxParams(absTable);
    console.log(`  Найдено параметров: ${Object.keys(tableParams).length}`);

    console.log('Чтение конфигурации ...');
    const configParams = parseXlsxParams(absConfig);
    console.log(`  Найдено параметров: ${Object.keys(configParams).length}`);

    console.log('Сравнение ...');
    const results = compare(tableParams, configParams, {});
    const report  = formatReport(results, absTable, absConfig, tableParams, configParams, null, 'Siemens');
    const absOut  = saveAndOpen(report, absTable, outputPath);
    console.log(`\nОтчёт сохранён: ${absOut}`);
    logSummary(results);
    return absOut;
}

// ─────────────────────────────────────────────────────────────────────────────
// Разбор аргументов CLI / интерактивный режим
// ─────────────────────────────────────────────────────────────────────────────

function parseArgs(argv) {
    const args = { word: null, xml: null, output: null, group: 1 };
    const pos  = [];
    for (let i = 0; i < argv.length; i++) {
        const a = argv[i];
        if (a === '-o' || a === '--output') { args.output = argv[++i]; }
        else if (a === '--group')           { args.group  = parseInt(argv[++i]) || 1; }
        else if (!a.startsWith('-'))        { pos.push(a); }
    }
    if (pos[0]) args.word = pos[0];
    if (pos[1]) args.xml  = pos[1];
    return args;
}

async function interactiveMode() {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const ask = (q) => new Promise(res => rl.question(q, res));

    console.log('='.repeat(62));
    console.log('  Сравнение таблицы уставок и файла конфигурации');
    console.log('='.repeat(62));
    console.log(`  Рабочая папка: ${process.cwd()}`);
    console.log('='.repeat(62));
    console.log();
    console.log('Введите пути к файлам (можно перетащить файл в окно консоли).');
    console.log();

    const modeStr = (await ask('Тип терминала: 1 = БЭ2704 (docx + xml), 2 = Siemens (xlsx + xml) [Enter = 1]: ')).trim();
    const isSiemens = modeStr === '2';
    console.log();

    let absOut;
    try {
        if (isSiemens) {
            const table = cleanPath(await ask('Таблица уставок (.xlsx): '));
            const xml   = cleanPath(await ask('Конфигурационный файл (.xml или .xrio): '));
            rl.close();
            console.log();
            // нормализуем: если пользователь передал .xrio — находим xml
            const xmlNorm = /\.xrio$/i.test(xml)
                ? xml.replace(/\.xrio$/i, '.xml')
                : xml;
            absOut = await runSiemens(table, xmlNorm, null);
        } else {
            const word   = cleanPath(await ask('Таблица уставок (.docx): '));
            const xml    = cleanPath(await ask('Файл конфигурации (.xml): '));
            const grpStr = (await ask('Группа уставок (1-4, Enter = 1): ')).trim();
            const group  = parseInt(grpStr) || 1;
            rl.close();
            console.log();
            absOut = await run(word, xml, null, group);
        }
    } catch (e) {
        console.error('\nОШИБКА:', e.message);
        writeErrorLog(e);
    }
    await pressEnterToExit();
}

// ─────────────────────────────────────────────────────────────────────────────
// Точка входа + перехват необработанных ошибок
// ─────────────────────────────────────────────────────────────────────────────

process.on('uncaughtException', async (err) => {
    console.error('\nНЕОЖИДАННАЯ ОШИБКА:', err.message);
    writeErrorLog(err);
    await pressEnterToExit().catch(() => {});
    process.exit(1);
});

(async () => {
    const args = parseArgs(process.argv.slice(2));
    if (args.word && args.xml) {
        try {
            const w = args.word.toLowerCase(), x = args.xml.toLowerCase();
            // Siemens: таблица.xlsx + конфиг.xml (или .xrio)
            const isSiemens = w.endsWith('.xlsx') && (x.endsWith('.xml') || x.endsWith('.xrio'));
            if (isSiemens) {
                const xmlNorm = x.endsWith('.xrio')
                    ? args.xml.replace(/\.xrio$/i, '.xml')
                    : args.xml;
                await runSiemens(args.word, xmlNorm, args.output);
            } else {
                await run(args.word, args.xml, args.output, args.group);
            }
        } catch (e) {
            console.error('\nОШИБКА:', e.message);
            writeErrorLog(e);
            process.exit(1);
        }
    } else {
        await interactiveMode();
    }
})();
