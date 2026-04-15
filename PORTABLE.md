# Portable build

## Goal
Собрать переносимый набор, который можно передать на Windows ПК без установки Python и без интернета.

## Build
```bat
build_portable.bat
```

Альтернатива по шагам:

```bat
pyinstaller --onefile --name digsi_readable.exe digsi_readable.py
python tools\build_portable.py --output-dir portable
```

## Output
- `portable/digsi_readable_portable/`
- `portable/digsi_readable-portable.zip`

В комплект включаются (если найдены):
- `digsi_readable.exe`
- `pipeline.exe`
- `compare_tool.exe`
- `SiemensPie/`
- `run_digsi_readable.bat`
- `portable_manifest.json` (контрольные суммы)

## Usage on target PC
```bat
run_digsi_readable.bat <input_config> -o out --report
```

