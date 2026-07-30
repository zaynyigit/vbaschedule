# Günlük Proje Çizelgesi — Şartname

**Sürüm 1.0 · 29 Temmuz 2026 · Yiğit Bayır**

Hedef ortam: **Windows Excel**, makro etkin tek dosya (`.xlsm`).
Tasarım kararları [brainstorm oturumunda](../.superpowers/brainstorm/) mockup'lar
üzerinden onaylandı; bu belge onların yazılı karşılığıdır.

---

## 1. Kapsam ve temel kararlar

| Konu | Karar | Gerekçe |
|---|---|---|
| Arayüz tekniği | Her şey çalışma sayfasına çizilir (UserForm yok) | Modern görünüm; ayrıca baskı bedava gelir |
| Modaller | Kendi sayfasına çizilen tam ekran panel, tek motor | Altı ayrı panel yerine tek yerde hata ayıklama |
| Kullanım | Tek kişilik | Excel makrolu dosyada eş zamanlı düzenleme veri bozar |
| Zaman modeli | Görev **gerçek tarih** tutar, ay görünümü kesişimi çizer | Aylar arası sarkan iş tek kayıt olarak kalır |
| Veri | Aynı kitapta gizli sayfalar (`xlSheetVeryHidden`) | Tek dosya, harici bağımlılık yok |
| Güvenlik | Erişim düzeni, kasa değil | VBA'da gerçek kriptografi yok |

### Kapsam dışı
- Eş zamanlı çok kullanıcılı düzenleme
- Gerçek kimlik doğrulama / şifreleme
- Mac Excel desteği (ActiveX parola kutusu çalışmaz)
- Görevler arası bağımlılık okları

---

## 2. Veri modeli

Her veri sayfasının 1. satırı başlık, veri 2. satırdan başlar.
Tüm okuma/yazma `modVeri` üzerinden geçer; arayüz kodu gizli sayfalara
doğrudan dokunmaz.

### `_Kullanicilar`
`ID · AdSoyad · KullaniciAdi · Parola · Rol · Foto · YGorevEkle · YYorum ·
YGecikme · YKullaniciYonet · Sayfa · Aktif · Renk`

- `Rol` ∈ {`admin`, `kullanici`}
- Dört ayrı yetki bayrağı; `admin` hepsini otomatik taşır
- `Renk` avatar rengi (hex, `#` yok)
- Parola `ParolaKaristir()` ile saklanır — **geri döndürülebilir değildir ama
  kriptografik de değildir**

### `_Projeler` / `_Kategoriler`
`ID · Ad · Renk · Sira · Aktif`

Bir proje birden çok görev taşır. Kategori görevin rozet rengini belirler.

### `_Gorevler`
`ID · ProjeID · KategoriID · Sira · Ad · Sorumlular · OncekiAyNot ·
PlanBaslangic · PlanBitis · GerceklesenBitis · Durum · CubukYazisi · Yuzde ·
Aciklama · OlusturanID · OlusturmaTarihi`

- `Sorumlular` virgülle ayrılmış kullanıcı ID listesi (çok sorumlu)
- `GerceklesenBitis` boşsa iş devam ediyor
- `Durum` ∈ {`plan`, `devam`, `tamam`}

### `_Yorumlar`
`ID · GorevID · KullaniciID · Tarih · Metin`

### `_Gecikmeler`
`ID · GorevID · Neden · Sorumluluk · KaydedenID · Tarih` — görev başına tek kayıt

### `_OffGunler`
`Tarih · Aciklama` — idari tatil işaretli günler

### `_Ayarlar`
`Anahtar · Deger` — aktif ay/yıl, oturumdaki kullanıcı, ID sayaçları

---

## 3. Ekran yerleşimi

Sayfa: `Cizelge`. Kılavuz çizgileri ve satır/sütun başlıkları kapalı.

```
satır 2   Uygulama başlığı + araç çubuğu (‹ Önceki | Bugün | Sonraki › | + Görev |
          Kullanıcılar | Kategoriler | Yazdır/PDF | Çıkış)
satır 4   Renk açıklama şeridi
satır 6   ┌──────────────── TEMMUZ / 2026 (koyu bant, 31 sütuna birleşik) ────────┐
satır 7   │ gün harfleri  P S Ç P C C P …                                         │
satır 8   │ gün numaraları 1 2 3 4 …                                              │
satır 9+  veri satırları
```

| Sütun | İçerik | Genişlik |
|---|---|---|
| A | Proje renk şeridi | 5 pt |
| B | Kategori (renkli rozet) | 62 pt |
| C | Sıra no | 22 pt |
| D | Sorumlu avatarları | 66 pt |
| E | Görev adı | 150 pt |
| F | Önceki aydan kalan | 118 pt |
| G…AK | Ayın günleri (1–31) | 17 pt |
| AL | İlerleme (halka + %) | 56 pt |
| AM | Açıklama | 150 pt |

**Bölme dondurma:** `G9` — sol blok ve takvim başlığı sabit kalır.
28/29/30 günlük aylarda fazla gün sütunları gizlenir; sağdaki iki sütun
sabit konumda durur.

### Satır tipleri
- **Proje grup satırı** — açık zemin, proje adı + görev sayısı, süre bandı, proje ortalama halkası
- **Görev satırı** — kategori rozeti, sıra, avatarlar, ad, önceki not, bar, halka, açıklama

---

## 4. Görsel dil

### Durum renkleri
| Durum | Bar |
|---|---|
| Tamamlandı | Dolu yeşil `#16A34A` |
| Devam ediyor | Dolu sarı `#F59E0B` |
| Planlandı | Saydam gri `#94A3B8` (%78 saydam), iki ucunda kısa dik gri çizgi |

Bar içine `CubukYazisi` yazılır (Tooling, Malzeme alımı…) — bar 40 pt'den darsa
yazı düşürülür.

### Gün zeminleri
Hafta sonu `#F3F6FA` · İdari tatil `#E6EAF1` · Bugün `#FEF2F2`
Bugün ayrıca yukarıdan aşağı **kırmızı dikey çizgi** (`#DC2626`, 1.75 pt).

### İlerleme halkası
Üç şekilden oluşur: arka plan çemberi → `msoShapePie` ilerleme dilimi
(saat 12'den saat yönünde) → ortada zemin renginde oyuk. Yüzde hücreye yazılır.
`%100`'de pie yerine tam daire çizilir (sıfır açı sorununu önlemek için).

### Avatarlar
En fazla 3 daire yan yana, fazlası `+N`. Baş harfler, kullanıcı rengi zemin.
Hücre açıklamasında tam sorumlu listesi (üzerine gelince çıkar).

### Yorum rozeti
Barın sağ ucunda koyu daire, içinde yorum sayısı. Tıklanınca detay paneli açılır.

---

## 5. Gecikme takibi

Görev iki bitiş tarihi taşır: **planlanan** ve **gerçekleşen**.

```
gecikme = gerçekleşen − planlanan            (iş bittiyse)
gecikme = bugün − planlanan                   (iş bitmedi ve planı geçtiyse)
gecikme = 0                                   (diğer hâllerde)
```

Gecikme varsa çizelgeye üç öge eklenir:

1. Planlanan bitiş gününde **beyaz dolgulu, kırmızı çemberli nokta**
2. Oradan gerçekleşen bitişe **kırmızı kesikli çizgi** (`msoLineDashDot`)
3. Ucunda dik kırmızı kapak ve `+N g` etiketi

Şeride tıklandığında **kırmızı başlıklı gecikme paneli** açılır — normal yorum
panelinden bilerek farklı. İçinde planlanan/gerçekleşen/sapma üçlüsü, neden
metni (zorunlu), sorumluluk seçimi (İç ekip / Dış yüklenici / Tedarikçi /
Müşteri / Onay süreci / Diğer) ve konuşma geçmişi.

Görev kaydedilirken gecikme oluştuysa ve daha önce kayıt yoksa panel
kendiliğinden açılır.

---

## 6. Yetkiler

| Yetki | Ne yapar |
|---|---|
| `GorevEkle` | Görev ekleyebilir / düzenleyebilir |
| `Yorum` | Yorum yazabilir |
| `Gecikme` | Gecikme nedeni kaydedebilir |
| `KullaniciYonet` | Kullanıcı ve kategori yönetebilir |

`admin` rolü dördünü de otomatik taşır. Her kullanıcının erişebileceği sayfa
ayrıca seçilir.

**Bu güvenlik değildir.** Tek kişilik kullanımda yanlışlıkla veri bozmayı
engelleyen bir çittir. Sayfa gizleme ve VBA parolası internetteki araçlarla
dakikalar içinde aşılır.

---

## 7. Etkileşim haritası

| Eylem | Sonuç |
|---|---|
| Gün hücresine çift tık | O tarihte görev ekleme paneli |
| Görev adına çift tık | Görev düzenleme paneli |
| Gün numarasına çift tık | İdari tatil işaretle / kaldır |
| Bara veya yorum rozetine tık | Görev detayı + yorumlar |
| Kırmızı şeride tık | Gecikme paneli |
| Görev satırına sağ tık | Düzenle / Yorumlar / Gecikme / Sil menüsü |

---

## 8. Çıktı

`modBaski` sayfa düzenini kurar:

- Baskı alanı: başlıktan son görev satırına
- Yön, sığdırma (1 sayfa / genişlik 1 sayfa / ölçeksiz)
- Başlık satırları ve sol blok her sayfada tekrarlanır
- Alt bilgi: uygulama adı, ay, telif, sayfa no
- İstenirse gecikme nedenleri ve yorumlar ek sayfa olarak basılır
- `ExportAsFixedFormat` ile PDF

**A1/A2 kısıtı:** Excel'in hazır kağıt sabitleri A4 ve A3'tür. A2/A1 için kod
Windows kağıt kodlarını (66/67) dener; yazıcı sürücüsü desteklemiyorsa geri
okuma ile anlar, A3 yatay tek sayfaya düşer ve kullanıcıyı iki seçenekle
uyarır (A1 destekleyen sürücü kur, ya da vektörel A3 PDF'i matbaada büyüt).

---

## 9. Modül haritası

| Modül | Sorumluluk | Satır |
|---|---|---|
| `modSabitler` | Sabitler, yerleşim, renk dönüşümü (`Hx`) | 193 |
| `modVeri` | Veri katmanı — tüm okuma/yazma | 558 |
| `modKurulum` | Sayfaları ve örnek veriyi kurar | 360 |
| `modCizelge` | Izgarayı çizer (tek giriş: `Ciz`) | 840 |
| `modModal` | Genel panel motoru | 366 |
| `modGiris` | Giriş ekranı, oturum | 289 |
| `modGezinme` | Ay geçişleri, sayfa olayları | 78 |
| `modGorev` | Görev ekle/düzenle/sil | 273 |
| `modYorum` | Detay ve yorum paneli | 135 |
| `modGecikme` | Gecikme paneli | 93 |
| `modKullanici` | Kullanıcı ve kategori yönetimi | 194 |
| `modBaski` | Sayfa düzeni ve PDF | 285 |

Ayrıca sayfa kodu olarak yapıştırılan iki dosya: `ThisWorkbook` ve `Cizelge`.

### Tasarım kuralları
- Renkler **asla** çıplak sayı olarak yazılmaz; her zaman `Hx("RRGGBB")`.
  VBA renkleri BGR sıralı sakladığı için elle çevirmek hata kaynağıdır.
- Çizim şekilleri `gx_` önekiyle adlandırılır ve her çizimde silinir.
  Kalıcı arayüz düğmeleri `UI_` önekini kullanır, silinmez.
- `modCizelge.Ciz` tek giriş noktasıdır; veri değişen her yerden çağrılır.

---

## 10. Doğrulama durumu

Bu makinede Excel bulunmadığı için kod **hiç çalıştırılmadı**. Statik olarak
doğrulananlar:

- ✅ Modüller arası tüm çağrılar çözülüyor (0 kırık referans)
- ✅ Public isim çakışması yok
- ✅ Tüm `OnAction` hedefleri mevcut ve Public
- ✅ Sub/Function/If/For/With/Select/Do blok dengeleri tamam
- ✅ Modül içi tekrar eden yordam adı yok
- ✅ Renk sabitleri hex→BGR dönüşümü koda bırakıldı (16 sabitin 14'ü elle
  yazıldığında yanlıştı; sınıf olarak ortadan kaldırıldı)
- ❌ Çalışma zamanı davranışı — Windows Excel'de ilk çalıştırmada doğrulanacak

İlk çalıştırmada hata çıkarsa hata numarası, hangi işlemde çıktığı ve
`Hata Ayıkla` ile sarıya boyanan satır yeterli.
