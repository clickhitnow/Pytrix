# Flask ile Excel Veri Analizi Uygulaması

Bu uygulama, yüklenen Excel dosyalarını işleyerek çeşitli özet raporlar üreten ve sonuçları kullanıcıya sunan bir Flask tabanlı web uygulamasıdır.

## İçerik

- [Proje Hakkında](#proje-hakkında)
- [Özellikler](#özellikler)
- [Kurulum ve Çalıştırma](#kurulum-ve-çalıştırma)
- [Kullanılan Teknolojiler](#kullanılan-teknolojiler)
- [Uygulama Yapısı](#uygulama-yapısı)
- [Lisans](#lisans)

---

## Proje Hakkında

Bu proje, yüklenen Excel dosyalarını analiz ederek, çeşitli özet tablolar oluşturur. Bu tablolar aşağıdaki bilgileri içerir:

- **Departmanlara iletilen taleplerin ay bazında dağılımı**
- **Mesai içinde ve mesai dışında iletilen talepler**
- **En çok talep ileten müşterilerin listeleri**
- **Talebin geliş yolu analizi (Mail, CRM, GSM, Zoom, WhatsApp vb.)**

---

## Özellikler

- **Excel dosyası yükleme ve önizleme** özelliği
- **Dinamik pivot tablolarla raporlama**
- **Raporların Excel formatında indirilmesi**

## Teknolojiler

- Flask
- Pandas
- XlsxWriter

## Kurulum ve Çalıştırma

### Gereksinimler

Uygulamayı çalıştırmak için aşağıdaki bağımlılıkları yükleyin:

```bash
pip install flask pandas openpyxl xlsxwriter

