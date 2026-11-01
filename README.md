# Al-Yamama RFI Workspace

English data-entry application based on the supplied Excel master workbook.

## التشغيل مباشرة من GitHub

[![Open in GitHub Codespaces](https://github.com/codespaces/badge.svg)](https://codespaces.new/Haedr0Ahmaed/Yamama?quickstart=1)

1. اضغط **Open in GitHub Codespaces** ثم **Create codespace** باستخدام حسابك.
2. انتظر اكتمال الإعداد؛ يتم تثبيت المتطلبات وتشغيل التطبيق تلقائياً.
3. افتح تبويب **Ports** واضغط رابط المنفذ **8000** إن لم تفتح الواجهة تلقائياً.

تبقى الواجهة باللغة الإنجليزية، لغة ملف Excel. تبقى السجلات داخل مساحة Codespace الحالية؛ صدّر نسخة Excel أو احتفظ بنسخة من `data/entries.json` قبل حذفها. قد تخضع Codespaces لحصص حساب GitHub أو تكلفته.

GitHub Pages لا يشغّل خادم Python؛ استخدم Codespaces لتشغيل التطبيق من المتصفح.

## Run locally

```sh
git clone https://github.com/Haedr0Ahmaed/Yamama.git
cd Yamama
python3 -m pip install -r requirements.txt
python3 server.py
```

Open http://localhost:8000 in your local browser.

## Behavior

Only the first two worksheets are used: **RFI Form** is the output template and **log** supplies drawing and activity options and the next RFI number. Other worksheets are excluded from exports.

- Discipline: **CIV**; inspection time: **12:00**; ITP item: **1**.
- Inspection, submission and receipt dates use calendar pickers, allowing new dates.
- Submitted by: **\*\*** / **Haedr Ahmaed** / **Dhiya Mahdi**. One submission date repeats across all three columns.
- QC receipt: the two subcontractor name cells remain blank; Al-Yamama QC is **Haedr Ahmaed**, Sinoma QC is **Engr. Ga Zang Dong Zhi**. One receipt date repeats across all four columns.
- **Area / Location** is free text. **Drawing No. & Rev No.** and **Inspection Activity / Description** are closed dropdowns from the log.
- Subcontractor and ITR number repeat the current RFI template values. The remaining form sections retain their template values, with blank signatures.
- Saved requests can be searched and edited. Each saved row has an **RFI Excel** download that fills the RFI Form for that record. The main export uses the latest saved record.
- Export contains **RFI Form** and **log** only; saved requests are appended to the log. The original template remains unchanged.
- Fixed values are enforced by the server as well as displayed read-only.

Data persists in `data/entries.json`; keep a backup before deleting a Codespace. This is a single-workspace application without authentication, intended for local use or private Codespaces ports.

## Validation

```sh
python3 -m unittest discover -s tests -v
```

Checks cover fixed values, calendar date validation, drawing choices, populated form cells, shared dates and the two-sheet Excel export.
