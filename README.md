# Photo Sorter

Turn a messy folder of photos into a tidy archive. Folders are named after **the year each photo was taken** and **the camera or phone that took it**, read from EXIF data.

**Before**
```
Photos\
├── IMG_0001.jpg
├── IMG_4521.jpg
├── DSC_0193.jpg
└── scan_01.jpg
```

**After**
```
Photos\
├── 2021\
│   └── Canon EOS 80D\
│       └── IMG_0001.jpg
├── 2023\
│   ├── Apple iPhone 13 Pro\
│   │   └── IMG_4521.jpg
│   └── NIKON CORPORATION NIKON D750\
│       └── DSC_0193.jpg
└── _No EXIF\
    └── scan_01.jpg
```

## Features

- Uses the date the photo was **taken** (`DateTimeOriginal`), not the date the file was modified
- Adds the brand to model names that leave it out (`SM-G991B` → `samsung SM-G991B`)
- Never overwrites files: duplicate names get `_1`, `_2`, ...
- Photos without EXIF data go into `_No EXIF`
- **Nothing to install.** It uses Windows' built-in WIA library.

## Usage

- Drag & drop a photo folder onto `photo-sorter.vbs`, **or**
- double-click `photo-sorter.vbs` and pick a folder.

Only the photos directly inside the chosen folder are processed. Subfolders are left alone.

> ⚠️ Photos are **moved**, not copied. Try it on a copy of your folder first.

## Settings

Edit the top of the script:

```vb
Const BY_YEAR   = True          ' create a year folder level
Const BY_CAMERA = True          ' create a camera model folder level
Const NO_EXIF_FOLDER = "_No EXIF"
```

For example, set `BY_CAMERA = False` to sort by year only.

## Supported files

`.jpg` and `.jpeg`

## License

MIT
