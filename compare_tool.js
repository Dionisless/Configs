#!/usr/bin/env node
/**
 * Инструмент сравнения таблицы уставок (Word .docx) и файла конфигурации (XML).
 *
 * Использование:
 *   compare_tool.exe <таблица.docx> <конфиг.xml> [-o отчёт.txt] [--group N]
 *   compare_tool.exe          (интерактивный режим — запросит пути файлов)
 */

'use strict';

const fs   = require('fs');
const path = require('path');
const readline = require('readline');

const AdmZip = require('adm-zip');
const { XMLParser } = require('fast-xml-parser');

// ─────────────────────────────────────────────────────────────────────────────
// Утилиты
// ─────────────────────────────────────────────────────────────────────────────

const ID_RE = /\[(\d{6})\]/;

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
    // "-", "–", "—" и "не используется" приравниваются к пустому
    if (/^[-–—]+$/.test(s) || /^не\s+используется$/i.test(s)) return '';
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

function formatReport(results, wordPath, xmlPath, wordParams, xmlParams, group) {
    const SEP = '='.repeat(72);
    const lines = [];
    const add = (s = '') => lines.push(s);

    const now = new Date().toLocaleString('ru-RU');
    add(SEP);
    add('ОТЧЁТ СРАВНЕНИЯ ТАБЛИЦЫ УСТАВОК И ФАЙЛА КОНФИГУРАЦИИ');
    add(SEP);
    add(`Сформирован:          ${now}`);
    add(`Таблица уставок:      ${wordPath}`);
    add(`Файл конфигурации:    ${xmlPath}`);
    add(`Группа уставок XML:   ${group}`);
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

    // 3. Только в таблице
    add(SEP);
    add(`3. ЕСТЬ В ТАБЛИЦЕ, НЕТ В КОНФИГЕ — ${ow.length} позиций`);
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
    add(`4. ЕСТЬ В КОНФИГЕ (со значением), НЕТ В ТАБЛИЦЕ — ${ox.length} позиций`);
    add(SEP);
    if (ox.length) {
        for (const p of ox) {
            add(`  ${p.name}`);
            if (p.path) add(`    Раздел: ${p.path}`);
            const vStr = p.value + (p.unit ? ` ${p.unit}` : '');
            add(`    Значение в конфиге: ${vStr}`);
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
            add(`    Значение в конфиге: ${vStr}`);
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

async function run(wordPath, xmlPath, outputPath, group) {
    // Привести к абсолютным путям
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

    // Имя выходного файла — рядом с таблицей, только ASCII в имени
    const date = new Date().toISOString().slice(0, 10);
    const absOut = outputPath
        ? path.resolve(outputPath)
        : path.join(path.dirname(absWord), `report_${date}.txt`);

    // Записать с BOM для корректного отображения в Блокноте Windows
    fs.writeFileSync(absOut, '\uFEFF' + report, 'utf8');

    const d = results.differences.length;
    const s = (results.smallDiff || []).length;
    const w = results.onlyWord.length;
    const x = results.onlyXml.length;
    const e = results.tableEmpty.length;
    console.log(`\nОтчёт сохранён: ${absOut}`);
    console.log(`Итог: Различий: ${d}  |  <1%: ${s}  |  Только в таблице: ${w}  |  Только в конфиге: ${x}  |  Пусто в таблице: ${e}`);

    // На Windows — открыть отчёт в Блокноте автоматически
    if (process.platform === 'win32') {
        try {
            require('child_process').spawn('notepad.exe', [absOut], { detached: true, stdio: 'ignore' }).unref();
        } catch (_) { /* ignore if notepad not available */ }
    }

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

    const word   = cleanPath(await ask('Таблица уставок (.docx): '));
    const xml    = cleanPath(await ask('Файл конфигурации (.xml): '));
    const grpStr = (await ask('Группа уставок (1-4, Enter = 1): ')).trim();
    const group  = parseInt(grpStr) || 1;
    rl.close();

    console.log();
    try {
        await run(word, xml, null, group);
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
            await run(args.word, args.xml, args.output, args.group);
        } catch (e) {
            console.error('\nОШИБКА:', e.message);
            writeErrorLog(e);
            process.exit(1);
        }
    } else {
        await interactiveMode();
    }
})();
