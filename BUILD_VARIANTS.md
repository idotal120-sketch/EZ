# ניהול שתי גרסאות — מלאה וקלה

## גישה מומלצת: **קוד אחד, שני קבצי spec**

- **קובץ אחד** של קוד (ענף `main`).
- **שני build**:
  - `AutoLang.spec` → `dist\AutoLang.exe` — גרסה מלאה (דיבור, תרגום אופליין, כל התכונות).
  - `AutoLangLight.spec` → `dist\AutoLangLight.exe` — גרסה קלה (בלי דיבור, בלי Argos; תרגום אונליין עובד).

## פקודות

```bash
# גרסה מלאה
pyinstaller --clean AutoLang.spec

# גרסה קלה (build ו־הפעלה מהירים יותר)
pyinstaller --clean AutoLangLight.spec
```

## מה מוצאים מהגרסה הקלה

| מוצאים | סיבה |
|--------|------|
| `faster_whisper`, `ctranslate2`, `sentencepiece` | דיבור → הכי כבד (מודלים + DLL) |
| `argostranslate` | תרגום אופליין |
| `speech_module` | תלוי ב־faster_whisper |
| `pyaudio` | הקלטה לדיבור |

הגרסה הקלה עדיין כוללת: תיקון שפה, איות, דקדוק (API), תרגום אונליין (Google), תיבת הקלדה, Undo.

## Git

- **לא חובה** ענף נפרד לגרסה קלה — מספיק שני ה־spec באותו ענף.
- אם תרצה להפריד: ענף `light` עם רק `AutoLangLight.spec` ועדכונים מ־`main` (merge/cherry-pick).

## טיפ

אם תוסיף בעתיד תכונה כבדה נוספת — הוסף אותה ל־`AutoLang.spec` והוצא אותה (או את התלויות שלה) ב־`AutoLangLight.spec` וב־`excludes`.
