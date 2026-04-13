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

/** Нормализация значения для сравнения */
function normalizeValue(val) {
    if (val === null || val === undefined) return '';
    let s = String(val).trim().replace(/[\n\r\t]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (!s) return '';
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
            params[nid] = {
                id: nid, fullId,
                name: s['@_Name'] || `[${nid}]`,
                value: String(value),
                unit: s['@_Unit'] || '',
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

    return params;
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
                };
            }
        }
    }

    return params;
}

// ─────────────────────────────────────────────────────────────────────────────
// Сравнение
// ─────────────────────────────────────────────────────────────────────────────

function compare(wordParams, xmlParams) {
    const allIds = [...new Set([...Object.keys(wordParams), ...Object.keys(xmlParams)])].sort();
    const differences = [], onlyWord = [], onlyXml = [];

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
            const xn = normalizeValue(xp.value);
            if (!wn && !xn) continue;
            if (wn !== xn) {
                differences.push({
                    id: nid, name: xp.name, path: xp.path || '',
                    valueWord: wp.value, unitWord: wp.unit || '',
                    valueXml: xp.value,  unitXml: xp.unit  || '',
                    primaryValues: wp.primaryValues || false,
                    allXmlValues: xp.allValues || {},
                });
            }
        }
    }

    return { differences, onlyWord, onlyXml };
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

    // 1. Различия
    const diffs = results.differences;
    add(SEP);
    add(`1. РАЗЛИЧИЯ ЗНАЧЕНИЙ — ${diffs.length} позиций`);
    add(SEP);
    if (diffs.length) {
        for (const d of diffs) {
            add(`  ${d.name}`);
            if (d.path) add(`    Раздел: ${d.path}`);
            const wStr = d.valueWord + (d.unitWord ? ` ${d.unitWord}` : '');
            const xStr = d.valueXml  + (d.unitXml  ? ` ${d.unitXml}`  : '');
            add(`    Таблица:  ${wStr || '(пусто)'}`);
            add(`    Конфиг:   ${xStr || '(пусто)'}`);
            if (d.primaryValues) add(`    * Значение в таблице указано в первичных единицах`);
            const av = d.allXmlValues;
            const avKeys = Object.keys(av);
            if (avKeys.length > 1) {
                add(`    Конфиг (все группы): ${avKeys.map(g => `Гр${g}=${av[g]}`).join('  ')}`);
            }
            add();
        }
    } else {
        add('  Различий нет.');
        add();
    }

    // 2. Только в таблице
    const ow = results.onlyWord;
    add(SEP);
    add(`2. ЕСТЬ В ТАБЛИЦЕ, НЕТ В КОНФИГЕ — ${ow.length} позиций`);
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

    // 3. Только в конфиге
    const ox = results.onlyXml;
    add(SEP);
    add(`3. ЕСТЬ В КОНФИГЕ (со значением), НЕТ В ТАБЛИЦЕ — ${ox.length} позиций`);
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

    add(SEP);
    add('КОНЕЦ ОТЧЁТА');
    add(SEP);
    return lines.join('\r\n');
}

// ─────────────────────────────────────────────────────────────────────────────
// Главная функция
// ─────────────────────────────────────────────────────────────────────────────

async function run(wordPath, xmlPath, outputPath, group) {
    console.log(`Чтение таблицы уставок: ${wordPath} ...`);
    const wordParams = parseDocxTable(wordPath);
    console.log(`  Найдено параметров: ${Object.keys(wordParams).length}`);

    console.log(`Чтение конфигурации:   ${xmlPath} ...`);
    const xmlParams = parseXmlConfig(xmlPath, group);
    console.log(`  Найдено параметров: ${Object.keys(xmlParams).length}`);

    console.log('Сравнение ...');
    const results = compare(wordParams, xmlParams);
    const report  = formatReport(results, wordPath, xmlPath, wordParams, xmlParams, group);

    const out = outputPath || path.join(
        path.dirname(wordPath),
        'отчет_сравнения_' + new Date().toISOString().slice(0, 10) + '.txt'
    );

    // Записать с BOM для корректного отображения в Блокноте Windows
    fs.writeFileSync(out, '\uFEFF' + report, 'utf8');

    console.log(`\nОтчёт сохранён: ${out}`);
    const d = results.differences.length;
    const w = results.onlyWord.length;
    const x = results.onlyXml.length;
    console.log(`Итог: Различий: ${d}  |  Только в таблице: ${w}  |  Только в конфиге: ${x}`);

    return out;
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

    console.log('='.repeat(60));
    console.log(' Сравнение таблицы уставок и файла конфигурации');
    console.log('='.repeat(60));
    console.log();

    const word   = (await ask('Путь к таблице уставок (.docx): ')).trim();
    const xml    = (await ask('Путь к файлу конфигурации (.xml): ')).trim();
    const grpStr = (await ask('Группа уставок (1-4, Enter = 1): ')).trim();
    const group  = parseInt(grpStr) || 1;
    const output = (await ask('Файл отчёта (Enter = авто): ')).trim() || null;
    rl.close();

    console.log();
    try {
        await run(word, xml, output, group);
    } catch (e) {
        console.error('ОШИБКА:', e.message);
        process.exitCode = 1;
    }
}

(async () => {
    const args = parseArgs(process.argv.slice(2));
    if (args.word && args.xml) {
        try {
            await run(args.word, args.xml, args.output, args.group);
        } catch (e) {
            console.error('ОШИБКА:', e.message);
            process.exitCode = 1;
        }
    } else {
        await interactiveMode();
    }
})();
