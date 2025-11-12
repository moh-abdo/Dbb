````markdown name=access/README_ar.md
```markdown
# فتح قاعدة بيانات Access (CarRental.accdb)

هذه التعليمات لفتح ملف قاعدة بيانات Microsoft Access عبر سكربت بسيط. ضع ملف قاعدة البيانات `CarRental.accdb` في نفس المجلد مع السكربت ثم شغّل أحد الملفات التالية.

الملفات المضمّنة:

- `open_database.bat`  — ملف دفعي Windows (Batch) لفتح القاعدة باستخدام ربط الملفات الافتراضي.
- `open_database.vbs`  — سكربت VBScript يفتح القاعدة من نفس المجلد.
- `open_database.ps1`  — سكربت PowerShell لفتح القاعدة.

كيفية الاستخدام:

1. ضع `CarRental.accdb` في نفس المجلد مع السكربتات (مثال: مجلد `access/`).
2. لتشغيل الباتش: انقر مزدوجًا على `open_database.bat` أو افتحه من سطر الأوامر.
3. لتشغيل VBScript: انقر مزدوجًا على `open_database.vbs` أو شغّله باستخدام `cscript`/`wscript`.
4. لتشغيل PowerShell: افتح PowerShell وانتقل إلى المجلد ثم شغّل `.
 open_database.ps1`. قد تحتاج لتغيير سياسة التنفيذ مؤقتًا عبر:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

ملاحظات مهمة:

- يجب أن يكون Microsoft Access مثبتًا على الجهاز لفتح ملفات `.accdb`.
- إذا لم يكن ملف `CarRental.accdb` موجودًا في نفس المجلد، سيعرض السكربت رسالة خطأ.
- تستطيع إنشاء اختصار (shortcut) إلى أي واحد من السكربتات ووضعه على سطح المكتب لفتح القاعدة بسرعة.

إذا تريد أن أرفع ملف `CarRental.accdb` نموذجياً في المستودع (كمثال) أو أُدرج سكربت يحمّل القاعدة من رابط، أخبرني وسأضيف ذلك.
```