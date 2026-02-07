# راهنمای استفاده از سیستم مدیریت فونت

## نحوه اعمال فونت به فرم‌ها

### 1. برای فرم‌های جدید

در سازنده فرم خود، بعد از `InitializeComponent()` فراخوانی کنید:

```csharp
public FormYourForm()
{
    InitializeComponent();
    // اعمال فونت به فرم
    FontSettings.ApplyFontToForm(this);
}
```

### 2. برای فرم‌های موجود

در فرم‌های موجود که `InitializeComponent` دستی نوشته شده، این خط را اضافه کنید:

```csharp
public FormPersonnelRegister()
{
    InitializeComponent();
    FontSettings.ApplyFontToForm(this);  // این خط را اضافه کنید
    LoadProvinces();
    LoadOtherCombos();
    UpdateInconsistency();
}
```

### 3. استفاده از فونت‌های مختلف در کد

شما می‌توانید به جای ایجاد دستی فونت، از فونت‌های تنظیم‌شده استفاده کنید:

#### قبل: (استفاده قدیمی)
```csharp
Label lblTitle = new Label
{
    Text = "سیستم مدیریت پرسنل",
    Font = new Font("Tahoma", 20, FontStyle.Bold),  // قدیمی
    // ...
};
```

#### بعد: (استفاده جدید)
```csharp
Label lblTitle = new Label
{
    Text = "سیستم مدیریت پرسنل",
    Font = FontSettings.TitleFont,  // جدید - فونت عنوان
    // ...
};
```

### 4. فونت‌های موجود:

```csharp
FontSettings.TitleFont      // عنوان بزرگ (با سایز پایه + 8)
FontSettings.HeaderFont     // سرتیتر (با سایز پایه + 6)
FontSettings.ButtonFont     // دکمه‌ها (با سایز پایه, Bold)
FontSettings.LabelFont      // برچسب‌ها (با سایز پایه + 1)
FontSettings.TextBoxFont    // جعبه متن و ComboBox
FontSettings.BodyFont       // متن عادی
```

### 5. ایجاد فونت سفارشی:

```csharp
// فونت با سایز پایه
Font myFont = FontSettings.GetCustomFont();

// فونت با سایز بزرگتر (+4 سایز)
Font biggerFont = FontSettings.GetCustomFont(sizeOffset: 4);

// فونت Bold
Font boldFont = FontSettings.GetCustomFont(style: FontStyle.Bold);

// ترکیبی
Font customFont = FontSettings.GetCustomFont(sizeOffset: 2, style: FontStyle.Italic);
```

## فرم‌هایی که باید به‌روزرسانی شوند:

1. ✅ **MainForm.cs** - به‌روزرسانی شد
2. ❌ **FormPersonnelRegister.cs** - نیاز به به‌روزرسانی
3. ❌ **FormPersonnelEdit.cs** - نیاز به به‌روزرسانی
4. ❌ **FormPersonnelDelete.cs** - نیاز به به‌روزرسانی
5. ❌ **FormPersonnelSearch.cs** - نیاز به به‌روزرسانی
6. ❌ **FormPersonnelAnalytics.cs** - نیاز به به‌روزرسانی

## مثال کامل:

```csharp
using System;
using System.Drawing;
using System.Windows.Forms;

namespace PersonnelManagementApp
{
    public partial class FormExample : Form
    {
        public FormExample()
        {
            InitializeComponent();
            
            // اعمال خودکار فونت به تمام کنترل‌ها
            FontSettings.ApplyFontToForm(this);
        }

        private void InitializeComponent()
        {
            this.Text = "فرم نمونه";
            this.Size = new Size(600, 400);

            // عنوان
            Label lblTitle = new Label
            {
                Text = "عنوار فرم",
                Location = new Point(200, 20),
                Size = new Size(200, 40),
                Font = FontSettings.TitleFont  // استفاده از فونت عنوان
            };
            this.Controls.Add(lblTitle);

            // دکمه
            Button btnAction = new Button
            {
                Text = "ذخیره",
                Location = new Point(250, 100),
                Size = new Size(100, 40),
                Font = FontSettings.ButtonFont  // استفاده از فونت دکمه
            };
            this.Controls.Add(btnAction);
        }
    }
}
```

## نحوه استفاده از فرم تنظیمات:

1. برنامه را اجرا کنید
2. از منوی اصلی دکمه **"تنظیمات"** را بزنید
3. فونت دلخواه را انتخاب کنید (مثلاً B Nazanin)
4. اندازه فونت را تغییر دهید (8-24)
5. دکمه **"پیش‌نمایش"** را بزنید تا نتیجه را ببینید
6. دکمه **"ذخیره تنظیمات"** را بزنید
7. برنامه را ببندید و مجدداً باز کنید

تغییرات بر روی **تمام فرم‌ها** اعمال خواهد شد!

## فایل‌های مرتبط:

- `FontSettings.cs` - کلاس اصلی مدیریت فونت
- `FormSettings.cs` - فرم تنظیمات
- `FormFontApplier.cs` - کلاس کمکی برای اعمال فونت
- `FontConfig.xml` - فایل ذخیره تنظیمات (خودکار ایجاد می‌شود)

---

**نوت**: پس از تغییر تنظیمات فونت، حتماً برنامه را مجدداً راه‌اندازی کنید تا تغییرات اعمال شوند.