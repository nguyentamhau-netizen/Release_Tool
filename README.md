# Release Note Test Result Tool

Tool local de:

- chon loai release: `SIT`, `UAT`, `PROD`
- chon file Release Note `.xlsx` tu may
- chon sheet hop le trong file
- generate file Test Result `.xlsx` vao thu muc `output/generated`
- neu co cau hinh Taiga local thi tu dien `Status` va `QC PIC`

## Cai dat

```powershell
cd C:\Users\Archer\SC-NEW\release-note-test-result-tool
python -m venv .venv
.\.venv\Scripts\activate
pip install -e .
```

## Chay app

```powershell
python -m release_note_tool
```

Hoac:

```powershell
release-note-tool
```

## Rule hien tai

- File input la `.xlsx` local
- Tool tu quet cac sheet co du cot:
  `No`, `Sprint`, `Service`, `Module`, `Link`, `US ID`, `Type`, `US Name`
- Output tao trong `output/generated`
- Ten file theo format:
  `KD - Test Result - <TYPE> Request DD-MM-YYYY.xlsx`

## Taiga Config

Neu muon tu dien `Status` va `QC PIC`, tao file `taiga.local.json` o root repo.
Ban co the copy tu `taiga.local.example.json`.

Field:

- `baseUrl`
- `projectSlug`
- `username`
- `password`
- `qcNames`: danh sach ten QC hop le de giu lai

## Assumption hien tai

- Với `User Story`, tool lay `assigned_to`
- Với `Issue`, tool uu tien lay `watchers`, neu khong co thi fallback `assigned_to`
- Neu co `qcNames`, ten nao khong nam trong danh sach se bi loai khoi `QC PIC`
- Neu chua co `taiga.local.json`, `Status` va `QC PIC` se de trong
- Ngay trong ten file mac dinh la ngay hien tai, nhung co the sua truc tiep tren app
- Tool tu tim header row bang cach do cot `No` cung cac cot bat buoc khac
