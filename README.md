# Staj Belgesi Sesli Giriş ve Kontrol Araçları
# Internship Record Voice Entry & Checker

**[🇹🇷 Türkçe](#türkçe)** · **[🇬🇧 English](#english)**

Github Pages : https://salihayesilyurt.github.io/internship-records-toolkit/internship_voice_entry.html

---

## Türkçe

Bu repo, öğrencilerin staj belgelerini (staj defteri, gizli sicil fişi, puantaj cetveli) inceleyen öğretim elemanları için iki araç içerir:

1. **Sesli giriş uygulaması.** Belgelere bakarken klavyeye dokunmadan, konuşarak Excel tablosu doldurursunuz.
2. **Kontrol scripti.** Doldurulan Excel'deki hataları bulur: yanlış numara, aynı öğrencinin farklı yazılmış adı, tarih ve iş günü uyuşmazlığı, resmi tatiller, kabul kuralları, dağınık notlar. Bulduğu hataları raporlar, düzeltilebilenleri düzeltir.

Programlama bilmeniz gerekmez. Aşağıdaki adımları sırayla izlemeniz yeterli.

### İçindekiler
- [Repodaki dosyalar](#repodaki-dosyalar)
- [Excel tablosunun yapısı](#excel-tablosunun-yapısı)
- [1. Sesli giriş uygulaması](#1-sesli-giriş-uygulaması)
- [2. Kontrol scripti](#2-kontrol-scripti)
- [Önerilen iş akışı](#önerilen-iş-akışı)
- [Gizlilik](#gizlilik)

### Repodaki dosyalar

| Dosya | Ne işe yarar |
|---|---|
| `internship_voice_entry.html` | Sesli giriş uygulaması. Tarayıcıda açılır, kurulum gerektirmez. |
| `internship_checker.ipynb` | Kontrol scriptinin Google Colab sürümü. En kolay yol budur. |
| `internship_checker.py` | Aynı kontrol scriptinin bilgisayarda çalışan sürümü. |
| `README.md` | Bu kılavuz. |

### Excel tablosunun yapısı

İki araç da ilk sayfada (Sheet1) aşağıdaki 13 sütunu bekler. Başlıklar birebir aynı olmalıdır.

| Sütun | Örnek | Kural |
|---|---|---|
| SN | 1 | Sıra numarası |
| Öğrenci No | 2023141043 | 10 haneli |
| Ad-Soyad | Ahmet Yılmaz | Aynı öğrencinin her satırında aynı yazılmalı |
| Staj Yeri | Multinet Teknoloji A.Ş. | |
| Sorumlu Kişi | Mehmet Kaya | |
| Başlama Tarihi | 01.07.2026 | Tarih biçiminde |
| Bitiş Tarihi | 05.08.2026 | Başlamadan sonra olmalı |
| Gizli Sicil Fişi | Var / Yok | |
| Her Sayfa İmzalı mı | Evet / Hayır | |
| Puantaj Cetveli | Var / Yok | |
| Yapılan Gün | 25 | Belgedeki iş günü |
| Kabul Edilen Gün | 25 | İmza **Hayır**, puantaj **Yok** veya gizli sicil **Yok** ise **0** |
| Açıklama | Gizli sicilde fotoğraf yok. | Serbest not |

Cumartesi çalışan öğrencilerin açıklamasına **"Cumartesi çalışma dilekçesi var."** yazın. Kontrol scripti bu öğrencilerde cumartesileri de iş günü sayar.

---

### 1. Sesli giriş uygulaması

#### Gerekenler
- **Google Chrome** veya **Microsoft Edge**. Firefox ses tanımayı desteklemez, Safari'de düzgün çalışmayabilir.
- Mikrofon (dizüstü bilgisayarın kendi mikrofonu yeterli, kulaklık mikrofonu daha iyi sonuç verir).
- İnternet bağlantısı. Konuşmayı yazıya tarayıcı çevirir ve bunun için internet gerekir.

#### İndirme ve açma
1. GitHub'da `internship_voice_entry.html` dosyasına tıklayın.
2. Sağ üstteki **indirme** simgesine (⬇ "Download raw file") tıklayın.
3. İnen dosyaya çift tıklayın. Varsayılan tarayıcınız Chrome veya Edge değilse dosyaya sağ tıklayıp **Birlikte aç → Google Chrome** seçin.

#### Mikrofon iznini bir kez verip unutmak için (önerilir)
Dosyayı bilgisayardan çift tıklayarak açtığınızda (adres `file://` ile başlar) Chrome mikrofon iznini **kalıcı hatırlamaz** ve sayfayı her açışınızda bir kez sorar. Uygulamayı bir web adresinden açarsanız izin bir kez verilir ve bir daha sorulmaz. İki yol var:

- **GitHub Pages (en kolayı):** Repoda **Settings → Pages → Source: Deploy from a branch → Branch: main → Save** yapın. Birkaç dakika sonra uygulama `https://KULLANICI-ADINIZ.github.io/REPO-ADI/internship_voice_entry.html` adresinde açılır. Bu adresi yer imlerine ekleyin, ilk açılışta **İzin ver** deyin. Verileriniz yine yalnızca sizin tarayıcınızda kalır, GitHub'a gitmez.
- **Kendi bilgisayarınızda (internetsiz sunucu):** Python kuruluysa dosyanın bulunduğu klasörde komut istemine `python -m http.server 8000` yazın ve Chrome'da `http://localhost:8000/internship_voice_entry.html` adresini açın. Komut istemi açık kaldığı sürece çalışır.

#### Kullanım
1. Mevcut bir Excel'e devam edecekseniz **Excel aç** ile dosyanızı seçin. Kayıtlar yüklenir ve bir sonraki boş satırdan devam edilir. Sıfırdan başlıyorsanız bu adımı atlayın.
2. Ekrana bakmadan çalışmak istiyorsanız **🔊 Sesli okuma** menüsünden bir mod seçin (aşağıda açıklandı).
3. **🎤 Dinlemeyi başlat** düğmesine basın. Tarayıcı mikrofon izni isterse **İzin ver** deyin.
4. Belgeye bakarak değerleri sırayla söyleyin. Her değer aktif hücreye yazılır ve otomatik olarak bir sonraki alana geçilir. Örnek bir öğrenci:

   > "yeni satır" → "2023 141 043" → "Ahmet Yılmaz" → "Multinet Teknoloji" → "Mehmet Kaya" → "1 temmuz 2026" → "5 ağustos" → "var" → "evet" → "var" → "yirmi beş" → "aynı" → "gizli sicilde fotoğraf yok" → "yeni satır"

5. Bitince **Excel'e aktar** düğmesine basın. `STAJ_2026.xlsx` dosyası İndirilenler klasörünüze iner.

#### Sesli komutlar

| Söyleyin | Ne olur |
|---|---|
| **yeni satır** | Yeni öğrenci satırı açılır, Öğrenci No'dan başlanır |
| **geç** / **boş** | Alan boş bırakılır, sonrakine geçilir |
| **geri** | Önceki alana dönülür |
| **sil** | Aktif alan temizlenir |
| **aynı** / **aynen** | Üst satırdaki değer kopyalanır. Kabul Edilen Gün alanında Yapılan Gün kopyalanır. |
| **neredeyim** | Hangi satır ve alanda olduğunuz sesli söylenir |
| **satırı oku** | Aktif satırın tamamı sesli okunur |
| Alan adı: **numara, ad, kurum, sorumlu, başlama, bitiş, sicil, imza, puantaj, yapılan, kabul, açıklama** | O alana atlanır |
| **dur** | Dinleme durur |

#### Değerleri söyleme ipuçları
- **Tarih:** "1 temmuz 2026" ya da "14 ağustos". Yıl söylenmezse 2026 kabul edilir.
- **Sayılar:** "otuz iki" ya da "32", ikisi de olur. Öğrenci numarasını üçer dörder gruplar halinde söyleyebilirsiniz.
- **Var/Yok, Evet/Hayır:** Tek kelime yeter.
- **Açıklama:** Bu alanda söyledikleriniz arka arkaya eklenir ve alan kendiliğinden geçmez. Bitince "yeni satır" deyin.

#### Sesli okuma modları
| Mod | Ne yapar |
|---|---|
| Kapalı | Sadece kısa bip sesleri: ince bip kabul edildi, kalın bip anlaşılamadı demek |
| Alan adını oku | Her adımda sıradaki alanın adı söylenir ("Staj yeri") |
| Değeri tekrarla + alan adı | Önce anlaşılan değer, sonra sıradaki alan söylenir ("Ahmet Yılmaz. Staj yeri"). Yanlış duyulan değeri hemen fark edip "geri" diyebilirsiniz. |

Uygulama okurken mikrofon açık kalır, ancak bu sırada duyulanlar yok sayılır. Böylece uygulama kendi sesini komut sanmaz. Okuma sürerken düğme **mor** görünür; değeri okuma bittikten sonra söyleyin. Hoparlör yerine kulaklık kullanmak tanımayı daha da iyileştirir.

#### Verileriniz nerede?
- Girdiğiniz her şey **tarayıcınızda otomatik saklanır**. Sayfayı kapatıp açsanız da kayıtlar kaybolmaz.
- Tarayıcı geçmişini/çerezleri temizlerseniz bu kayıtlar silinir. Bu yüzden **ara ara Excel'e aktarın**.
- Yanlış yazılmış bir değeri tablodaki hücreye tıklayıp klavyeyle de düzeltebilirsiniz.
- **Son satırı sil** düğmesi en alttaki satırı siler.

#### Sorun giderme
| Sorun | Çözüm |
|---|---|
| "Bu tarayıcı ses tanımayı desteklemiyor" | Dosyayı Chrome veya Edge ile açın |
| Mikrofon çalışmıyor | Adres çubuğundaki 🔒 / mikrofon simgesine tıklayıp mikrofona izin verin, sayfayı yenileyin |
| Her açılışta mikrofon izni soruyor | Dosya bilgisayardan (`file://`) açılmış. GitHub Pages veya `localhost` adresinden açın ([yukarıya bakın](#mikrofon-iznini-bir-kez-verip-unutmak-için-önerilir)). |
| Dinleme arada kesiliyor | Uygulama kendiliğinden yeniden başlatır. Uzun sessizliklerden sonra bir kez "neredeyim" demek yeterli. |
| Sesli okuma Türkçe değil ya da hiç ses yok | Windows: **Ayarlar → Saat ve dil → Dil** bölümünden Türkçe dil paketini ve konuşma özelliğini yükleyin. Bilgisayarın sesinin açık olduğunu kontrol edin. |
| Bir değer sürekli yanlış anlaşılıyor | Daha yavaş ve tane tane söyleyin ya da hücreye tıklayıp klavyeyle yazın |

---

### 2. Kontrol scripti

Excel doldurulduktan sonra çalıştırın. Script şunları kontrol eder:

- **Öğrenci No:** 10 haneli mi? Numaradaki bölüm kodu diğerlerinden farklı mı?
- **Ad tutarlılığı:** Aynı numarada farklı ad (ör. "Tapşo" / "Tabşo"), aynı adda farklı numara var mı? Ad yerine yanlışlıkla "Aynen" gibi bir komut yazılmışsa aynı numaralı diğer satırdan düzeltir.
- **Tarihler:** Okunamayan veya bozuk tarih ("03.08.206" → 03.08.2026), başlama tarihinin bitişten sonra olması, aynı öğrencinin stajlarının çakışması.
- **İş günü:** Başlama ve bitiş arasındaki hafta içi günlerden resmi tatilleri çıkarır ve sonucu Yapılan Gün ile karşılaştırır. Tutmuyorsa doğru bitiş tarihini önerir. Cumartesiler sayıldığında tutuyorsa "cumartesi çalışmış olabilir" uyarısı verir.
- **Kabul kuralları:** İmza Hayır, puantaj Yok veya gizli sicil Yok olduğu halde Kabul 0 değilse; belgeler tamken Kabul 0 ise; Kabul, Yapılan Gün'den büyükse.
- **Açıklamalar:** Aynı anlama gelen notları tek biçime getirir ("Gizli silindi fotoğraf yok", "Gizli sicil'de fotoğraf yok", "fotoğraf eksik" → "Gizli sicilde fotoğraf yok."). Sesle yazarken kalan "yeni satış 2021" gibi artıkları siler.
- **Yazım:** "A.ş" → "A.Ş.", "Ltd Şti" → "Ltd. Şti.". Aynı kurumun veya kişinin farklı yazımlarını ayrıca listeler, ama kendiliğinden birleştirmez.

#### Yöntem A: Google Colab (önerilen, hiçbir şey kurmadan)
1. Google hesabınızla [colab.research.google.com](https://colab.research.google.com) adresini açın.
2. **Dosya → Not defterini aç → GitHub** sekmesine bu reponun adresini yapıştırın ve `internship_checker.ipynb` dosyasını seçin.
   (Ya da dosyayı GitHub'dan indirip **Dosya → Not defteri yükle** ile yükleyin.)
3. Menüden **Çalışma zamanı → Tümünü çalıştır** seçin.
4. En alttaki hücrede **Dosya seç** düğmesi çıkar. Excel dosyanızı seçin.
5. Birkaç saniye sonra `..._checked.xlsx` dosyası bilgisayarınıza iner.

#### Yöntem B: Kendi bilgisayarınızda
1. [python.org](https://www.python.org/downloads/) adresinden Python'u kurun. Windows'ta kurulumda **"Add Python to PATH"** kutusunu işaretleyin.
2. Komut istemini (Windows: Başlat → "cmd") açıp bir kez şunu yazın:
   ```
   pip install openpyxl
   ```
3. `internship_checker.py` ile Excel dosyanızı aynı klasöre koyun, o klasörde komut istemini açıp şunu çalıştırın:
   ```
   python internship_checker.py STAJ_2026.xlsx
   ```
4. Aynı klasörde `STAJ_2026_checked.xlsx` oluşur.

#### Çıktı dosyasını okuma
Orijinal dosyanız değişmez, sonuçlar yeni bir dosyaya yazılır.

| Sayfa | İçerik |
|---|---|
| **Sheet1** | Düzeltilmiş veri. **Sarı** hücre otomatik düzeltildi, **turuncu** hücreyi elle kontrol etmeniz gerekiyor. Hücrenin üzerine gelince yorumda ne yapıldığı veya ne önerildiği görünür. |
| **Kontrol Raporu** | Tüm bulgular. "Kontrol edin" olanlar en üstte. |
| **İş Günü Hesabı** | Her satırın hesap detayı: aradaki tatiller, hesaplanan gün, fark. |
| **Yazım Farkları** | Aynı kurum veya kişi olabilecek farklı yazımlar |

#### Her yıl yapılacak ayar
Dosyanın başındaki `HOLIDAYS` listesine yeni yılın tatillerini aynı biçimde ekleyin:
```python
date(2027, 7, 15): "15 Temmuz",
```
Diğer ayarlar:
- `EVE_IS_HOLIDAY = False`: Arefe günleri (yarım gün) iş günü sayılır. Tatil saymak için `True` yapın.
- `STUDENT_NO_DIGITS = 10`: Öğrenci numarasının hane sayısı.
- `COLUMNS = dict(...)`: Excel başlıklarınız farklıysa buradan değiştirin.

### Önerilen iş akışı
1. Belgeleri sesli giriş uygulamasıyla girin, ara ara **Excel'e aktar**ın.
2. Bitince Excel'i kontrol scriptinden geçirin.
3. Turuncu hücreleri belgeye bakarak düzeltin. Düzeltmeleri ister Excel'de, ister uygulamada (**Excel aç** ile yükleyip) yapın.
4. Scripti son bir kez çalıştırın. "Kontrol edin" maddesi kalmayana kadar tekrarlayın.

### Gizlilik
- **Öğrenci verisi içeren Excel dosyalarını GitHub'a yüklemeyin.** Repo herkese açıksa herkes görebilir.
- Sesli giriş uygulaması verileri yalnızca sizin tarayıcınızda saklar. Konuşmanın yazıya çevrilmesi ise tarayıcının (Google/Microsoft) sunucularında yapılır. Kurumunuzun veri politikası buna izin veriyor mu, kontrol edin.
- Kontrol scripti tamamen yerelde (ya da kendi Colab oturumunuzda) çalışır, veriyi hiçbir yere göndermez.

---

## English

This repository contains two tools for faculty members who review students' internship documents (internship logbook, confidential evaluation form, attendance sheet):

1. **Voice entry app.** Fill in the Excel table by speaking while you look at the documents, without touching the keyboard.
2. **Checker script.** Finds errors in the completed Excel file: wrong student numbers, differently spelled names of the same student, date and workday mismatches, public holidays, acceptance rules, inconsistent notes. It reports what it finds and fixes what it safely can.

No programming knowledge is needed. Just follow the steps below in order.

> **Note:** The app's interface and its voice commands are in **Turkish**, because the documents and the spreadsheet are Turkish. The English translation of each command is given below.

### Contents
- [Files in this repository](#files-in-this-repository)
- [Excel table layout](#excel-table-layout)
- [1. Voice entry app](#1-voice-entry-app)
- [2. Checker script](#2-checker-script)
- [Suggested workflow](#suggested-workflow)
- [Privacy](#privacy)

### Files in this repository

| File | Purpose |
|---|---|
| `internship_voice_entry.html` | Voice entry app. Opens in the browser, nothing to install. |
| `internship_checker.ipynb` | Google Colab version of the checker. The easiest option. |
| `internship_checker.py` | The same checker as a script for your own computer. |
| `README.md` | This guide. |

### Excel table layout

Both tools expect these 13 columns on the first sheet (Sheet1). The headers must match exactly.

| Column | Meaning | Example | Rule |
|---|---|---|---|
| SN | Row number | 1 | |
| Öğrenci No | Student number | 2023141043 | 10 digits |
| Ad-Soyad | Full name | Ahmet Yılmaz | Same spelling on every row of the same student |
| Staj Yeri | Company / institution | Multinet Teknoloji A.Ş. | |
| Sorumlu Kişi | Supervisor | Mehmet Kaya | |
| Başlama Tarihi | Start date | 01.07.2026 | Date format |
| Bitiş Tarihi | End date | 05.08.2026 | Must be after start |
| Gizli Sicil Fişi | Confidential evaluation form | Var / Yok (yes / no) | |
| Her Sayfa İmzalı mı | Every page signed? | Evet / Hayır (yes / no) | |
| Puantaj Cetveli | Attendance sheet | Var / Yok | |
| Yapılan Gün | Days worked | 25 | Workdays in the documents |
| Kabul Edilen Gün | Days accepted | 25 | **0** if signed = **Hayır**, attendance = **Yok** or evaluation form = **Yok** |
| Açıklama | Notes | Gizli sicilde fotoğraf yok. | Free text |

For students who worked on Saturdays, write **"Cumartesi çalışma dilekçesi var."** ("Saturday work petition on file") in the notes. The checker then counts Saturdays as workdays for that row.

---

### 1. Voice entry app

#### Requirements
- **Google Chrome** or **Microsoft Edge**. Firefox does not support speech recognition, and Safari may not work properly.
- A microphone (a laptop's built-in mic is fine, a headset mic works better).
- An internet connection. The browser converts speech to text online.

#### Download and open
1. On GitHub, click `internship_voice_entry.html`.
2. Click the **download** icon at the top right (⬇ "Download raw file").
3. Double-click the downloaded file. If your default browser is not Chrome or Edge, right-click the file and choose **Open with → Google Chrome**.

#### Grant microphone permission once (recommended)
When you open the file by double-clicking it (the address starts with `file://`), Chrome does **not remember** the microphone permission and asks once every time you open the page. If you open the app from a web address instead, you allow it once and are never asked again. Two ways:

- **GitHub Pages (easiest):** In the repository go to **Settings → Pages → Source: Deploy from a branch → Branch: main → Save**. After a few minutes the app is available at `https://YOUR-USERNAME.github.io/REPO-NAME/internship_voice_entry.html`. Bookmark it and click **Allow** the first time. Your data still stays only in your browser and is not sent to GitHub.
- **On your own computer (local server, no internet needed for this part):** If Python is installed, open a command prompt in the file's folder, run `python -m http.server 8000`, and open `http://localhost:8000/internship_voice_entry.html` in Chrome. It works as long as the command prompt stays open.

#### How to use
1. To continue an existing file, click **Excel aç** (Open Excel) and pick it. The records are loaded and entry continues from the next empty row. Skip this step if you are starting from scratch.
2. To work without looking at the screen, choose a mode in **🔊 Sesli okuma** (read aloud). The modes are explained below.
3. Click **🎤 Dinlemeyi başlat** (Start listening) and allow microphone access when the browser asks.
4. Read the values from the document in order. Each value goes into the active cell and the app moves to the next field automatically. One student, for example:

   > "yeni satır" → "2023 141 043" → "Ahmet Yılmaz" → "Multinet Teknoloji" → "Mehmet Kaya" → "1 temmuz 2026" → "5 ağustos" → "var" → "evet" → "var" → "yirmi beş" → "aynı" → "gizli sicilde fotoğraf yok" → "yeni satır"

5. When you are done, click **Excel'e aktar** (Export to Excel). `STAJ_2026.xlsx` is saved to your Downloads folder.

#### Voice commands

| Say | Meaning | What happens |
|---|---|---|
| **yeni satır** | new row | Starts a new student row at Student No |
| **geç** / **boş** | skip / empty | Leaves the field empty and moves on |
| **geri** | back | Goes to the previous field |
| **sil** | delete | Clears the active field |
| **aynı** / **aynen** | same | Copies the value from the row above. In Days accepted, copies Days worked. |
| **neredeyim** | where am I | Says the current row and field aloud |
| **satırı oku** | read the row | Reads the whole active row aloud |
| A field name: **numara, ad, kurum, sorumlu, başlama, bitiş, sicil, imza, puantaj, yapılan, kabul, açıklama** | number, name, company, supervisor, start, end, form, signature, attendance, worked, accepted, notes | Jumps to that field |
| **dur** | stop | Stops listening |

#### Tips for speaking values
- **Dates:** "1 temmuz 2026" or "14 ağustos" (Turkish month names). If no year is said, 2026 is used.
- **Numbers:** "otuz iki" or "32" both work. Student numbers can be said in groups of digits.
- **Var/Yok, Evet/Hayır:** One word is enough.
- **Notes:** In this field, everything you say is appended and the app does not move on. Say "yeni satır" when finished.

#### Read-aloud modes
| Mode | What it does |
|---|---|
| Kapalı (off) | Short beeps only: a high beep means accepted, a low beep means not understood |
| Alan adını oku (read field name) | Says the next field's name at every step ("Staj yeri") |
| Değeri tekrarla + alan adı (repeat value + field name) | Says the recognised value, then the next field ("Ahmet Yılmaz. Staj yeri"). You can say "geri" right away if something was misheard. |

While the app is speaking, the microphone stays on but whatever it hears is ignored, so the app doesn't take its own voice as a command. The button turns **purple** while it speaks; say the value after it finishes. A headset instead of speakers improves recognition further.

#### Where is my data?
- Everything you enter is **saved automatically in your browser**. Closing and reopening the page does not lose it.
- Clearing browser history or cookies deletes those records, so **export to Excel regularly**.
- You can also click any cell and fix a value with the keyboard.
- **Son satırı sil** (Delete last row) removes the bottom row.

#### Troubleshooting
| Problem | Fix |
|---|---|
| "Bu tarayıcı ses tanımayı desteklemiyor" (browser not supported) | Open the file in Chrome or Edge |
| Microphone not working | Click the 🔒 / microphone icon in the address bar, allow the microphone, reload the page |
| Asks for microphone permission every time | The file was opened from disk (`file://`). Open it from GitHub Pages or `localhost` ([see above](#grant-microphone-permission-once-recommended)). |
| Listening stops now and then | The app restarts it automatically. After a long silence, saying "neredeyim" once is enough. |
| Read-aloud is not Turkish, or there is no sound | Windows: install the Turkish language pack with speech under **Settings → Time & language → Language**. Also check that the computer's sound is on. |
| A value keeps being misheard | Speak more slowly and clearly, or click the cell and type it |

---

### 2. Checker script

Run it after the Excel file is filled in. It checks:

- **Student No:** Is it 10 digits? Is the department code inside the number different from the others?
- **Name consistency:** Different names under the same number (e.g. "Tapşo" / "Tabşo"), or the same name under different numbers. If a command such as "Aynen" was accidentally written instead of a name, it is fixed from another row with the same number.
- **Dates:** Unreadable or broken dates ("03.08.206" → 03.08.2026), start date after end date, and overlapping internships of the same student.
- **Workdays:** Counts weekdays between start and end, subtracts public holidays and compares the result with Days worked. On a mismatch it suggests the correct end date. If the count matches once Saturdays are included, it warns that the student probably worked Saturdays.
- **Acceptance rules:** Days accepted not 0 although signature = Hayır, attendance = Yok or evaluation form = Yok; Days accepted 0 although all documents are complete; Days accepted greater than Days worked.
- **Notes:** Brings notes with the same meaning into one wording ("Gizli silindi fotoğraf yok", "Gizli sicil'de fotoğraf yok", "fotoğraf eksik" → "Gizli sicilde fotoğraf yok."). Removes leftovers from voice entry such as "yeni satış 2021".
- **Spelling:** "A.ş" → "A.Ş.", "Ltd Şti" → "Ltd. Şti.". Different spellings that may be the same company or person are listed, but not merged automatically.

#### Option A: Google Colab (recommended, nothing to install)
1. Open [colab.research.google.com](https://colab.research.google.com) with your Google account.
2. Go to **File → Open notebook → GitHub**, paste this repository's address and select `internship_checker.ipynb`.
   (Or download the file from GitHub and use **File → Upload notebook**.)
3. Choose **Runtime → Run all**.
4. A **Choose files** button appears in the last cell. Select your Excel file.
5. A few seconds later, `..._checked.xlsx` downloads to your computer.

#### Option B: On your own computer
1. Install Python from [python.org](https://www.python.org/downloads/). On Windows, tick **"Add Python to PATH"** during setup.
2. Open a command prompt (Windows: Start → "cmd") and run this once:
   ```
   pip install openpyxl
   ```
3. Put `internship_checker.py` and your Excel file in the same folder, open a command prompt in that folder and run:
   ```
   python internship_checker.py STAJ_2026.xlsx
   ```
4. `STAJ_2026_checked.xlsx` appears in the same folder.

#### Reading the output file
Your original file is not changed; the results go into a new file.

| Sheet | Content |
|---|---|
| **Sheet1** | Corrected data. A **yellow** cell was fixed automatically; an **orange** cell needs your manual check. Hover over the cell to see what was done or suggested. |
| **Kontrol Raporu** (Check report) | All findings, with "Kontrol edin" (check) items at the top |
| **İş Günü Hesabı** (Workday calculation) | Per-row details: holidays in range, calculated days, difference |
| **Yazım Farkları** (Spelling differences) | Spellings that may refer to the same company or person |

#### Yearly setting
Add the new year's public holidays to the `HOLIDAYS` list at the top of the file, in the same format:
```python
date(2027, 7, 15): "15 Temmuz",
```
Other settings:
- `EVE_IS_HOLIDAY = False`: Holiday eves (half days) count as workdays. Set to `True` to treat them as holidays.
- `STUDENT_NO_DIGITS = 10`: Number of digits in a student number.
- `COLUMNS = dict(...)`: Change these if your Excel headers are different.

### Suggested workflow
1. Enter the documents with the voice entry app and **export to Excel** from time to time.
2. When finished, run the Excel file through the checker.
3. Fix the orange cells by checking the documents, either in Excel or in the app (load it with **Excel aç**).
4. Run the checker once more. Repeat until no "Kontrol edin" items remain.

### Privacy
- **Do not upload Excel files that contain student data to GitHub.** If the repository is public, anyone can see them.
- The voice entry app keeps data only in your own browser. Speech-to-text, however, is done on the browser vendor's (Google/Microsoft) servers. Check whether your institution's data policy allows this.
- The checker runs entirely on your computer (or in your own Colab session) and sends data nowhere.
