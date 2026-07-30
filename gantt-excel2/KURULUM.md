# Kurulum — 5 dakika

Windows Excel gerekiyor. Mac Excel'de ActiveX parola kutusu ve bazı şekil
özellikleri çalışmaz.

## 1. Boş bir kitap aç ve makro etkin kaydet

Excel → yeni boş çalışma kitabı → **Dosya > Farklı Kaydet** →
tür olarak **Excel Makro Etkin Çalışma Kitabı (*.xlsm)** seç.
Adı ne olursa olsun, ör. `Cizelge.xlsm`.

## 2. Makroları etkinleştir

**Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarları > Makro Ayarları**
→ "Tüm makroları etkinleştir" (ya da dosyayı güvenilir bir konuma koy).

## 3. Modülleri içe aktar

`Alt + F11` ile VBA düzenleyiciyi aç. Sol taraftaki ağaçta projeye sağ tıkla →
**Dosyayı İçe Aktar**. `src/` klasöründeki **12 adet `.bas`** dosyasını sırayla ekle:

```
modSabitler.bas
modVeri.bas
modKurulum.bas
modCizelge.bas
modModal.bas
modGiris.bas
modGezinme.bas
modGorev.bas
modYorum.bas
modGecikme.bas
modKullanici.bas
modBaski.bas
```

> Sıra önemli değil, VBA hepsini birlikte derler.

## 4. Sayfa kodlarını yapıştır

Bu ikisi `.bas` olarak aktarılamaz — sayfa olayları ancak sayfanın kendi kod
penceresinde çalışır.

- Ağaçta **ThisWorkbook**'a çift tıkla → `src/SayfaKodu_ThisWorkbook.txt`
  içeriğini yapıştır.
- **Kurulum makrosunu çalıştırdıktan sonra** ağaçta beliren **Cizelge** sayfasına
  çift tıkla → `src/SayfaKodu_Cizelge.txt` içeriğini yapıştır.

## 5. Kurulumu çalıştır

`Alt + F8` → **Kurulum** → Çalıştır.

Sayfalar, biçimler, örnek 3 proje / 7 görev / 7 kullanıcı otomatik oluşur.
Giriş ekranı açılır:

```
Kullanıcı adı : admin
Parola        : 1234
```

## 6. Kaydet

`Ctrl + S`. Tür yine **.xlsm** kalmalı.

---

# Kullanım

| Ne yapmak istiyorsun | Nasıl |
|---|---|
| Yeni görev eklemek | Bir gün hücresine **çift tıkla** (veya üstteki `+ Görev`) |
| Görevi düzenlemek | Görev adına **çift tıkla** |
| Yorum yazmak / okumak | Görevin **çubuğuna tıkla** |
| Gecikme nedeni girmek | Kırmızı kesikli **şeride tıkla** |
| Bir günü tatil yapmak | Gün numarasına **çift tıkla** |
| Görevi silmek | Görev satırına **sağ tıkla** → 4 |
| Ay değiştirmek | Üstteki `‹ Önceki ay` / `Sonraki ay ›` |
| Kullanıcı eklemek | `Kullanıcılar` düğmesi (yönetici yetkisi ister) |
| PDF almak | `Yazdır / PDF` düğmesi |

---

# Bilmen gereken sınırlar

**Bunlar Excel'in sınırları, kodun eksiği değil.** Baştan söylüyorum ki
karşılaştığında hata sanma.

### A1 / A2 kağıt boyutu
Excel'in hazır kağıt listesi A4 ve A3 ile biter. A2/A1'e geçebilmek için
bilgisayarda o boyutu sunan bir yazıcı sürücüsü (plotter ya da özel boyutlu PDF
yazıcısı) kurulu olmalı. Yoksa kod seni uyarır ve A3 yatay tek sayfaya sığdırır.
Çıktı vektörel olduğu için matbaada A1'e büyütmek kalite kaybettirmez.

### Sağdaki sütunlar donmuyor
Excel yalnızca üstten ve soldan bölme dondurur. `İlerleme` ve `Açıklama`
sütunları 31 günün arkasında olduğu için ekranda görmek için sağa kaydırman
gerekir. Baskıda hepsi tek sayfada, sorun yok.

### Güvenlik
Giriş ekranı ve roller bir **erişim düzenidir, kasa değildir.** Parolalar basit
bir karıştırmayla saklanır; VBA projesini ve gizli sayfaları açmak isteyen biri
internetteki bir araçla dakikalar içinde açar. Gerçek gizlilik gerektiren veriyi
bu dosyada tutma.

### Eş zamanlı kullanım
Dosya tek kişilik tasarlandı. İki kişi aynı anda açıp yazarsa Excel makrolu
dosyalarda eş zamanlı düzenlemeyi desteklemediği için veri kaybı olur.

---

# Sorun çıkarsa

Kod tek satır bile çalıştırılmadan teslim edildi (bu makinede Excel yok).
Statik olarak doğrulananlar: modüller arası tüm çağrılar çözülüyor, isim
çakışması yok, tüm `OnAction` hedefleri mevcut, blok dengeleri tamam.
Çalışma zamanında yine de hata çıkabilir.

Hata alırsan bana şunu getir:

1. **Hata numarası ve metni** (kutuda yazan)
2. **Hangi işlemde** çıktı (kurulum / giriş / görev ekleme / PDF...)
3. VBA'da `Hata Ayıkla` düğmesi çıkarsa **sarı ile işaretlenen satır**

Bu üçüyle sorunu doğrudan bulabilirim.

Kurulumu bozduğunu düşünürsen `Alt+F8` → **KurulumSifirla** her şeyi siler ve
baştan kurar (onay ister).
