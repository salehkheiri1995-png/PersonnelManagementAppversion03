# سیستم مدیریت تصاویر پرسنل - راهنمای کامل

## خلاصه تغییرات

برای اضافه کردن سیستم مدیریت عکس به برنامه، موارد زیر اضافه شده‌اند:

### 1. فایل‌های جدید ایجاد شده

#### `ImageHelper.cs`
کلاس کمکی برای مدیریت تصاویر پرسنل:
- ذخیره عکس با نام کد ملی
- بارگذاری، حذف و تغییر نام عکس
- باز کردن دیالوگ انتخاب عکس
- ایجاد تصویر پیش‌فرض

#### `DATABASE_UPDATE_PHOTO.md`
راهنمای به‌روزرسانی دیتابیس با دستورالعمل‌های دقیق

---

## نحوه استفاده در FormPersonnelRegister

### 1. اضافه کردن فیلدهای ضروری

در ابتدای کلاس:

```csharp
private PictureBox pbPhoto;
private string selectedPhotoPath = string.Empty;
```

### 2. اضافه کردن PictureBox به فرم

در متد `InitializeComponent()` بعد از بخش سرتیتر:

```csharp
// PictureBox برای نمایش عکس
int photoBoxSize = 200;
pbPhoto = new PictureBox
{
    Location = new Point((formWidth - photoBoxSize) / 2, yHeader),
    Size = new Size(photoBoxSize, photoBoxSize),
    BorderStyle = BorderStyle.FixedSingle,
    SizeMode = PictureBoxSizeMode.Zoom,
    BackColor = Color.White
};
pbPhoto.Image = ImageHelper.CreateDefaultImage(photoBoxSize, photoBoxSize);
ApplyRoundedCorners(pbPhoto, 15);
this.Controls.Add(pbPhoto);
yHeader += photoBoxSize + 10;

// دکمه‌های مدیریت عکس
int btnPhotoWidth = 95;
int btnPhotoSpacing = 5;
int totalPhotoButtonWidth = (btnPhotoWidth * 2) + btnPhotoSpacing;
int xPhotoButtonStart = (formWidth - totalPhotoButtonWidth) / 2;

// دکمه انتخاب عکس
Button btnSelectPhoto = new Button
{
    Text = "انتخاب عکس",
    Location = new Point(xPhotoButtonStart, yHeader),
    Size = new Size(btnPhotoWidth, 35),
    Font = FontSettings.ButtonFont,
    BackColor = Color.LightBlue,
    ForeColor = Color.White
};
ApplyRoundedCorners(btnSelectPhoto, 10);
btnSelectPhoto.Click += BtnSelectPhoto_Click;
this.Controls.Add(btnSelectPhoto);

// دکمه حذف عکس
Button btnRemovePhoto = new Button
{
    Text = "حذف عکس",
    Location = new Point(xPhotoButtonStart + btnPhotoWidth + btnPhotoSpacing, yHeader),
    Size = new Size(btnPhotoWidth, 35),
    Font = FontSettings.ButtonFont,
    BackColor = Color.LightCoral,
    ForeColor = Color.White
};
ApplyRoundedCorners(btnRemovePhoto, 10);
btnRemovePhoto.Click += BtnRemovePhoto_Click;
this.Controls.Add(btnRemovePhoto);
yHeader += 45;

// تنظیم yRight و yLeft
int yRight = yHeader + 10;
int yLeft = yHeader + 10;
```

### 3. اضافه کردن Event Handlers

```csharp
private void BtnSelectPhoto_Click(object sender, EventArgs e)
{
    string photoPath = ImageHelper.OpenImageDialog();
    if (!string.IsNullOrEmpty(photoPath))
    {
        selectedPhotoPath = photoPath;
        Image img = Image.FromFile(photoPath);
        ImageHelper.DrawImageInPictureBox(pbPhoto, img);
    }
}

private void BtnRemovePhoto_Click(object sender, EventArgs e)
{
    selectedPhotoPath = string.Empty;
    pbPhoto.Image = ImageHelper.CreateDefaultImage(pbPhoto.Width, pbPhoto.Height);
}
```

### 4. اصلاح متد ClearForm

اضافه کنید:

```csharp
selectedPhotoPath = string.Empty;
pbPhoto.Image = ImageHelper.CreateDefaultImage(pbPhoto.Width, pbPhoto.Height);
```

### 5. اصلاح متد BtnSave_Click

بعد از اجرای INSERT موفق:

```csharp
db.ExecuteNonQuery(query, parameters);

// ذخیره عکس اگر انتخاب شده است
if (!string.IsNullOrEmpty(selectedPhotoPath))
{
    string nationalID = txtNationalID.Text.Trim();
    ImageHelper.SaveImage(selectedPhotoPath, nationalID);
}

MessageBox.Show("پرسنل با موفقیت ثبت شد!", "موفقیت", MessageBoxButtons.OK, MessageBoxIcon.Information);
```

---

## نحوه استفاده در FormPersonnelEdit

### 1. اضافه کردن فیلدهای ضروری

در ابتدای کلاس:

```csharp
private PictureBox pbPhoto;
private string selectedPhotoPath = string.Empty;
private string currentNationalID = string.Empty;
```

### 2. اضافه کردن PictureBox به فرم

مشابه FormPersonnelRegister در متد `InitializeComponent()`

### 3. اضافه کردن Event Handlers

مشابه FormPersonnelRegister

### 4. اصلاح متد BtnLoad_Click

بعد از بارگذاری داده‌های پرسنل:

```csharp
currentNationalID = row["NationalID"] != DBNull.Value ? row["NationalID"].ToString() : "";
txtNationalID.Text = currentNationalID;

// بارگذاری عکس اگر وجود دارد
Image photo = ImageHelper.LoadImage(currentNationalID);
if (photo != null)
{
    ImageHelper.DrawImageInPictureBox(pbPhoto, photo);
    selectedPhotoPath = ImageHelper.GetImageFilePath(currentNationalID);
}
else
{
    pbPhoto.Image = ImageHelper.CreateDefaultImage(pbPhoto.Width, pbPhoto.Height);
    selectedPhotoPath = string.Empty;
}
```

### 5. اصلاح متد BtnUpdate_Click

قبل از اجرای UPDATE:

```csharp
// بررسی تغییر کد ملی
string oldNationalID = currentNationalID;
string newNationalID = txtNationalID.Text.Trim();
bool nationalIDChanged = oldNationalID != newNationalID;

try
{
    db.ExecuteNonQuery(query, parameters);

    // مدیریت تصویر
    if (!string.IsNullOrEmpty(selectedPhotoPath))
    {
        if (selectedPhotoPath.StartsWith(ImageHelper.ImagesFolderPath))
        {
            // عکس از سیستم است - اگر کد ملی تغییر کرده، نام عکس را تغییر بده
            if (nationalIDChanged)
            {
                ImageHelper.RenameImage(oldNationalID, newNationalID);
            }
        }
        else
        {
            // عکس جدید انتخاب شده - ذخیره کن
            ImageHelper.SaveImage(selectedPhotoPath, newNationalID);
        }
    }

    MessageBox.Show("رکورد پرسنل با موفقیت به‌روزرسانی شد!");
    this.Close();
}
catch (Exception ex)
{
    MessageBox.Show("خطا در به‌روزرسانی رکورد: " + ex.Message);
}
```

---

## تغییرات دیتابیس

### روش 1: از طریق Design View در Access

1. باز کردن فایل `MyDatabase.accdb`
2. جدول `Personnel` را در Design View باز کنید
3. در انتها فیلد جدید اضافه کنید:
   - **Field Name**: `PhotoPath`
   - **Data Type**: `Short Text`
   - **Field Size**: `255`
   - **Required**: `No`

### روش 2: با SQL Query

```sql
ALTER TABLE Personnel ADD COLUMN PhotoPath TEXT(255);
```

---

## ساختار پوشه تصاویر

```
[App Directory]
└── PersonnelImages/
    ├── 0123456789.jpg
    ├── 9876543210.jpg
    └── ...
```

**نکات:**
- پوشه به طور خودکار ایجاد می‌شود
- نام فایل‌ها برابر با کد ملی پرسنل است
- فرمت ذخیره: JPG با کیفیت 90%

---

## مزایای این سیستم

✅ **ساده و کارآمد**: مدیریت آسان تصاویر بدون پیچیدگی  
✅ **یکپارچه**: هماهنگ با کد ملی پرسنل  
✅ **انعطاف‌پذیر**: پشتیبانی از فرمت‌های مختلف تصویر  
✅ **خودکار**: مدیریت خودکار نام‌گذاری و حذف  
✅ **بهینه**: ذخیره با کیفیت بهینه  

---

## رفع مشکلات احتمالی

### عکس نمایش داده نمی‌شود
- بررسی کنید پوشه `PersonnelImages` وجود دارد
- بررسی کنید فایل با نام کد ملی موجود است
- بررسی کنید مجوز خواندن فایل را دارید

### عکس ذخیره نمی‌شود
- بررسی کنید مجوز نوشتن در پوشه برنامه را دارید
- بررسی کنید فضای کافی در درایو وجود دارد

### خطای Database بعد از به‌روزرسانی
- بررسی کنید فیلد `PhotoPath` به جدول اضافه شده است
- بررسی کنید نسخه پشتیبان دیتابیس را دارید

---

## تست کامل سیستم

### چک‌لیست تست:

- [ ] ثبت پرسنل جدید با عکس
- [ ] ثبت پرسنل جدید بدون عکس
- [ ] بارگذاری پرسنل موجود با عکس
- [ ] بارگذاری پرسنل موجود بدون عکس
- [ ] تغییر عکس پرسنل
- [ ] حذف عکس پرسنل
- [ ] تغییر کد ملی (باید نام فایل عکس هم تغییر کند)
- [ ] انتخاب فرمت‌های مختلف عکس (JPG, PNG, BMP)
- [ ] نمایش تصویر پیش‌فرض برای پرسنل بدون عکس

---

**نویسنده:** سیستم AI پرپلکسیتی  
**تاریخ:** 2026-02-06  
**نسخه:** 1.0.0