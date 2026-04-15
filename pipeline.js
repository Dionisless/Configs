#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const isPkg = typeof process.pkg !== 'undefined';
const baseDir = isPkg ? path.dirname(process.execPath) : __dirname;
const siemensDir = path.join(baseDir, 'SiemensPie');

function fail(message) {
  console.error(`Ошибка: ${message}`);
  process.exit(1);
}

function usage() {
  console.log('Использование:');
  console.log('  pipeline.exe <input_par_or_prot_or_xml_or_xrio> <reference.xlsx> [-o report.txt]');
  console.log('');
  console.log('Примеры:');
  console.log('  pipeline.exe PAR.PSD "Таблица уставок.xlsx" -o report.txt');
  console.log('  pipeline.exe config.xml "Таблица уставок.xlsx" -o report.txt');
}

function runOrFail(command, args, options = {}) {
  const result = spawnSync(command, args, {
    stdio: 'inherit',
    windowsHide: true,
    ...options,
  });

  if (result.error) {
    fail(`не удалось запустить команду: ${command} (${result.error.message})`);
  }

  if (result.status !== 0) {
    fail(`команда завершилась с кодом ${result.status}: ${command}`);
  }
}

function parseArgs(argv) {
  if (argv.length === 0 || argv.includes('-h') || argv.includes('--help')) {
    usage();
    process.exit(0);
  }

  if (argv.length < 2) {
    usage();
    fail('недостаточно аргументов');
  }

  const input = path.resolve(argv[0]);
  const referenceXlsx = path.resolve(argv[1]);
  let outputReport = path.resolve(process.cwd(), 'report.txt');

  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    if ((arg === '-o' || arg === '--output') && argv[i + 1]) {
      outputReport = path.resolve(argv[i + 1]);
      i += 1;
      continue;
    }
    fail(`неизвестный аргумент: ${arg}`);
  }

  if (!fs.existsSync(input)) fail(`входной файл не найден: ${input}`);
  if (!fs.existsSync(referenceXlsx)) fail(`референсный xlsx не найден: ${referenceXlsx}`);

  return { input, referenceXlsx, outputReport };
}

function resolveXmlXrioFromInput(inputPath) {
  const ext = path.extname(inputPath).toLowerCase();

  if (ext === '.xml') {
    return {
      xmlPath: inputPath,
      xrioPath: inputPath.replace(/\.xml$/i, '.xrio'),
    };
  }

  if (ext === '.xrio') {
    return {
      xmlPath: inputPath.replace(/\.xrio$/i, '.xml'),
      xrioPath: inputPath,
    };
  }

  if (ext !== '.par' && ext !== '.prot' && ext !== '.psd' && ext !== '.pds' && ext !== '.psp' && ext !== '.dps') {
    fail(`неподдерживаемый формат входного файла: ${ext}`);
  }

  const converterCandidates = [
    process.env.SIEMENS_CONVERTER_CMD,
    path.join(baseDir, 'par_prot_converter.exe'),
    path.join(baseDir, 'par_prot_converter.js'),
    path.join(siemensDir, 'par_prot_converter.exe'),
    path.join(siemensDir, 'par_prot_converter.js'),
  ].filter(Boolean);

  const converter = converterCandidates.find((candidate) => fs.existsSync(candidate));
  if (!converter) {
    fail('конвертер PAR/PROT → XML/XRio не найден. Укажите путь через SIEMENS_CONVERTER_CMD или положите par_prot_converter.exe рядом с pipeline.exe');
  }

  const inputDir = path.dirname(inputPath);
  const baseName = path.basename(inputPath, path.extname(inputPath));
  const xmlPath = path.join(inputDir, `${baseName}.xml`);
  const xrioPath = path.join(inputDir, `${baseName}.xrio`);

  console.log(`[1/3] Конвертация ${path.basename(inputPath)} -> XML/XRio...`);
  if (converter.endsWith('.js')) {
    runOrFail(process.execPath, [converter, inputPath, '--out-dir', inputDir]);
  } else {
    runOrFail(converter, [inputPath, '--out-dir', inputDir]);
  }

  return { xmlPath, xrioPath };
}

function runSiemensPie(xmlPath, xrioPath) {
  const spExe = path.join(siemensDir, 'sp.exe');
  if (!fs.existsSync(spExe)) {
    fail(`sp.exe не найден: ${spExe}`);
  }

  const args = fs.existsSync(xrioPath) ? [xmlPath, xrioPath] : [xmlPath];
  console.log(`[2/3] Генерация xlsx через SiemensPie (${args.map((x) => path.basename(x)).join(', ')})...`);

  runOrFail(spExe, args, { cwd: siemensDir });

  const outXlsx = xmlPath.replace(/\.xml$/i, '.xlsx');
  if (!fs.existsSync(outXlsx)) {
    fail(`SiemensPie не создал файл: ${outXlsx}`);
  }

  return outXlsx;
}

function runCompareTool(referenceXlsx, xmlPath, outputReport) {
  const compareJs = path.join(baseDir, 'compare_tool.js');
  const compareExe = path.join(baseDir, 'compare_tool.exe');

  console.log('[3/3] Формирование отчёта compare_tool...');

  if (isPkg && fs.existsSync(compareExe)) {
    runOrFail(compareExe, [referenceXlsx, xmlPath, '-o', outputReport]);
    return;
  }

  if (!fs.existsSync(compareJs)) {
    fail(`compare_tool.js не найден: ${compareJs}`);
  }

  runOrFail(process.execPath, [compareJs, referenceXlsx, xmlPath, '-o', outputReport]);
}

function main() {
  const { input, referenceXlsx, outputReport } = parseArgs(process.argv.slice(2));

  const { xmlPath, xrioPath } = resolveXmlXrioFromInput(input);
  if (!fs.existsSync(xmlPath)) fail(`XML-файл не найден: ${xmlPath}`);
  if (!fs.existsSync(xrioPath)) {
    console.warn(`Предупреждение: XRio-файл не найден: ${xrioPath}. SiemensPie будет запущен только с XML.`);
  }

  const generatedXlsx = runSiemensPie(xmlPath, xrioPath);
  runCompareTool(referenceXlsx, xmlPath, outputReport);

  console.log('');
  console.log('Пайплайн завершён успешно.');
  console.log(`XML:        ${xmlPath}`);
  console.log(`XRio:       ${xrioPath}`);
  console.log(`SiemensPie: ${generatedXlsx}`);
  console.log(`Отчёт:      ${outputReport}`);
}

main();
