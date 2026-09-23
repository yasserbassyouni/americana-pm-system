Smart PM V1.3.4 - Continuous Running Hours

التعديل الرئيسي:
- عند استيراد Annual Excel لسنة مثل 2026، يبدأ عداد Running Hours تلقائياً من 01/01/2026.
- موعد PM لا يبدأ في Week 1 بشكل تلقائي.
- الموعد يحسب من: Interval Hours / Expected Running Hours per Week.
- الحساب يستخدم CEIL حتى لا يتم جدولة PM قبل الوصول فعلياً للساعات المطلوبة.
- 10000 hr عند 120 hr/week = 84 أسبوع تقريباً، أول موعد من 01/01/2026 يصبح تقريباً 12/08/2027.
- عند تنفيذ PM فعلياً، يتم حفظ Execution Date كنقطة بدء جديدة لعداد Running Hours للسنوات القادمة.
- عند بناء سنة جديدة، يستمر العداد من السنة السابقة ولا يبدأ من الصفر.
- Year Planner يعرض: First Due Date / Occurrences / Interval Weeks / PM Code.
- PM Code هو كود اللون/فئة PM في Excel (1..5)، وليس عدد الأسابيع.
