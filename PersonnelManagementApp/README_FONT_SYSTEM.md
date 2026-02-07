# سیستم مدیریت فونت - Personnel Management App

## خلاصه

سیستم مدیریت فونت متمرکزی که به شما امکان می‌دهد نوع فونت و اندازه فونت را در تمام فرم‌های برنامه به صورت متمرکز مدیریت کنید.

## ویژگی‌ها

✅ **تغییر فونت** - انتخاب از 12 فونت فارسی مختلف  
✅ **تغییر سایز** - اندازه فونت قابل تنظیم از 8 تا 24  
✅ **پیش‌نمایش زنده** - مشاهده تغییرات قبل از ذخیره  
✅ **ذخیره خودکار** - تنظیمات در فایل XML ذخیره می‌شود  
✅ **اعمال خودکار** - تمام فرم‌ها به صورت خودکار به‌روزرسانی می‌شوند  

## فایل‌های سیستم

| فایل | توضیحات |
|------|------------|
| `FontSettings.cs` | کلاس اصلی مدیریت فونت - ذخیره، بارگذاری و اعمال فونت |
| `FormSettings.cs` | فرم تنظیمات - رابط کاربری برای تغییر فونت |
| `FormFontApplier.cs` | کلاس کمکی - متدهای آسان برای اعمال فونت |
| `FontConfig.xml` | فایل تنظیمات - خودکار ایجاد می‌شود |

## فونت‌های پشتیبانی شده

```
1. Tahoma (پیش‌فرض)
2. B Nazanin
3. B Titr
4. B Lotus
5. B Zar
6. IRANSans
7. Yekan
8. Mitra
9. Arial
10. Vazir
11. Samim
12. Sahel
```

## نحوه استفاده

### 1. برای کاربر نهایی

```
1. برنامه را اجرا کنید
2. در منوی اصلی، دکمه "تنظیمات" را بزنید
3. فونت دلخواه را انتخاب کنید
4. اندازه فونت را تنظیم کنید (8-24)
5. "پیش‌نمایش" را بزنید تا نتیجه را ببینید
6. "ذخیره تنظیمات" را بزنید
7. برنامه را مجدداً راه‌اندازی کنید
```

### 2. برای توسعه‌دهندگان

#### اضافه کردن پشتیبانی فونت به فرم جدید:

```csharp
using PersonnelManagementApp;

public class MyNewForm : Form
{
    public MyNewForm()
    {
        InitializeComponent();
        
        // فقط این یک خط!
        FontSettings.ApplyFontToForm(this);
    }
}
```

#### استفاده از فونت‌های مختلف:

```csharp
// فونت‌های از پیش تعریف شده
Label title = new Label { Font = FontSettings.TitleFont };
Button btn = new Button { Font = FontSettings.ButtonFont };
TextBox txt = new TextBox { Font = FontSettings.TextBoxFont };
Label lbl = new Label { Font = FontSettings.LabelFont };

// فونت سفارشی
Font custom = FontSettings.GetCustomFont(sizeOffset: 2, style: FontStyle.Bold);
```

## معماری سیستم

```
╭─────────────────────╮
│  FontConfig.xml    │  (فایل تنظیمات)
╰──────────┬─────────╯
           │
           ↓ Load/Save
╭──────────┴──────────╮
│  FontSettings.cs   │  (مدیریت متمرکز)
╰────────┬───────────╯
         │
         ├─────────────────────────╮
         │                            │
         ↓                            ↓
╭───────────────╮      ╭─────────────────╮
│ FormSettings  │      │  All Forms      │
│ (رابط کاربری) │      │  (اعمال فونت)  │
╰───────────────╯      ╰─────────────────╯
```

## API عمومی

### FontSettings

```csharp
// فونت‌های از پیش تعریف شده
public static Font TitleFont      // عنوان بزرگ
public static Font HeaderFont     // سرتیتر
public static Font BodyFont       // متن عادی
public static Font ButtonFont     // دکمه‌ها
public static Font LabelFont      // برچسب‌ها
public static Font TextBoxFont    // جعبه متن

// تنظیمات پایه
public static string DefaultFontName  // نام فونت فعلی
public static float DefaultFontSize   // سایز فونت فعلی

// متدها
public static void SaveSettings(string fontName, float fontSize)
public static Font GetCustomFont(float sizeOffset = 0, FontStyle style = FontStyle.Regular)
public static string[] GetPersianFonts()
public static void ApplyFontToForm(Form form)
```

## مثال‌های کاربردی

### 1. تغییر فونت به B Nazanin با سایز 14

```csharp
FontSettings.SaveSettings("B Nazanin", 14f);
```

### 2. ایجاد عنوان بزرگ

```csharp
Label bigTitle = new Label
{
    Text = "عنوان اصلی",
    Font = FontSettings.GetCustomFont(sizeOffset: 10, style: FontStyle.Bold)
};
```

### 3. اعمال فونت به فرم جدید

```csharp
public class FormReport : Form
{
    public FormReport()
    {
        InitializeComponent();
        FontSettings.ApplyFontToForm(this);
    }
}
```

## رفع مشکلات

### فونت اعمال نمی‌شود

✅ مطمئن شوید `FontSettings.ApplyFontToForm(this)` در سازنده فرم فراخوانی شده  
✅ مطمئن شوید برنامه را مجدداً راه‌اندازی کرده‌اید  
✅ فایل FontConfig.xml را حذف کرده و دوباره امتحان کنید  

### خطای کامپایل

✅ مطمئن شوید `using System.Drawing;` اضافه شده  
✅ Rebuild Solution را اجرا کنید  
✅ مطمئن شوید FontSettings.cs در پروژه وجود دارد  

### فونت بعد از راه‌اندازی مجدد بر می‌گردد

✅ این عملکرد عادی است - تنظیمات بعد از راه‌اندازی مجدد اعمال می‌شود  

## فرم‌های پشتیبانی شده

- ✅ MainForm.cs
- ✅ FormSettings.cs
- ❌ FormPersonnelRegister.cs (نیاز به به‌روزرسانی دستی)
- ❌ FormPersonnelEdit.cs (نیاز به به‌روزرسانی دستی)
- ❌ FormPersonnelDelete.cs (نیاز به به‌روزرسانی دستی)
- ❌ FormPersonnelSearch.cs (نیاز به به‌روزرسانی دستی)
- ❌ FormPersonnelAnalytics.cs (نیاز به به‌روزرسانی دستی)

برای به‌روزرسانی فرم‌های باقی‌مانده، به فایل `UPDATE_ALL_FORMS.txt` مراجعه کنید.

## مدارک

- `FontConfig.xml` - فایل تنظیمات XML
- `FONT_USAGE_GUIDE.md` - راهنمای کامل استفاده
- `UPDATE_ALL_FORMS.txt` - دستورالعمل به‌روزرسانی

## نسخه

نسخه 1.0.0 - 2026-02-01

## توسعه‌دهنده

Saleh Kheiri - [salehkheiri1995@gmail.com](mailto:salehkheiri1995@gmail.com)

## مجوز

MIT License - برای پروژه PersonnelManagementApp

---

**توجه**: برای اعمال تغییرات فونت، حتماً برنامه را مجدداً راه‌اندازی کنید!