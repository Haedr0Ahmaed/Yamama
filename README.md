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

- 42 log fields organized into request details, intervention, readiness, QC and closure.
- One distinct source value is treated as fixed and repeated automatically. Blank cells do not count as values. RFI numbers increment from the master log.
- Variable fields are closed dropdowns populated from the source log; only AREA / LOCATION is free text. Existing source spelling is retained. Date and time choices are also restricted to source values.
- Hold/action, inspection result and rejection reason choices come from the original RFI form.
- Empty source fields remain disabled until a source option is available.
- Saved requests can be searched and edited. Data persists in data/entries.json.
- Excel export adds saved requests after the final numbered source record in the log, retaining the original nine worksheets. The original template is never modified. The RFI Form sheet retains its original contents; it is not filled per new request.
- This is a single-workspace application without user authentication. It is intended for local use, not public deployment.

The original workbook is included at data/template.xlsx. Back up data/entries.json to preserve entered records.
