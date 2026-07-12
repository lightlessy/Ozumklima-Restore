# Özüm Klima UTF-8 Geçiş Planı

## Amaç

Classic ASP uygulamasını, HTML çıktısını, form girişlerini ve veritabanı içeriklerini tek bir Unicode akışında çalıştırmak.

Hedef standart:

- ASP dosyaları: UTF-8 (tercihen BOM'suz; sunucuda sorun görülürse kontrollü olarak BOM'lu test edilir)
- ASP çalışma zamanı: CodePage 65001
- HTTP yanıtı: `text/html; charset=utf-8`
- HTML: `<meta charset="utf-8">`
- Form kabulü: UTF-8
- XML/JSON çıktıları: UTF-8
- Access `.mdb`: mevcut Text/Memo alanları Unicode metin tutar; kolonları “UTF-8 tipine” dönüştürmek gerekmez

## Kritik uyarı

Veritabanındaki metinlere toplu `Replace` uygulanmamalıdır. `ş -> þ`, `ı -> ý`, `ğ -> ð` gibi görüntüler iki farklı nedenle oluşabilir:

1. Veri veritabanında doğru Unicode'dur, yalnızca yanlış response charset ile gönderiliyordur.
2. Veri daha önce yanlış encoding ile kaydedilmiş ve gerçekten mojibake olmuştur.

Birinci durumda veriyi dönüştürmek mevcut doğru içeriği bozar. Bu nedenle önce audit, sonra yalnızca kanıtlanmış bozuk kayıtların dönüşümü yapılır.

## Aşama 0 — Yedek ve geri dönüş

1. Canlı `.mdb` dosyasının tarihli kopyasını alın.
2. Web root ve IIS ayarlarının yedeğini alın.
3. Migration sırasında yazma işlemlerini kısa süreli durdurun.
4. Geçişi önce staging/kopya veritabanında test edin.

Önerilen adlandırma:

- `TRvd678577-before-utf8-YYYYMMDD-HHMM.mdb`
- `TRvd678577-utf8-test.mdb`

## Aşama 1 — Veri audit'i

Önce `tools/utf8-audit.vbs` çalıştırılır.

Audit şu karakterleri ve örüntüleri arar:

- `ý`, `þ`, `ð`, `Ý`, `Þ`, `Ð`
- `Ã`, `Â`, `Ä`, `Å`, `â€`
- `&#...;` şeklinde gereksiz entity olarak saklanmış içerikler

Her şüpheli kayıt tablo, kolon ve primary key ile raporlanır. Dönüşüm yapılmaz.

## Aşama 2 — Runtime UTF-8 hazırlığı

Yeni `charset-utf8.asp` include'u staging üzerinde denenir:

```asp
<%@ Language="VBScript" CodePage="65001" %>
<%
Response.Buffer = True
Response.CodePage = 65001
Session.CodePage = 65001
Session.LCID = 1055
Response.Charset = "utf-8"
%>
```

Not: ASP directive dosyanın ilk satırında olmalıdır. Include dosyasındaki directive bazı IIS kurulumlarında geçersiz sayılabilir. En güvenli yöntem, public giriş sayfalarının en başına directive eklemek; include içinde yalnızca Response/Session ayarlarını tutmaktır.

## Aşama 3 — HTML başlıklarının standardizasyonu

Tüm public ve panel sayfalarında:

Silinecek:

```html
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
```

Eklenecek:

```html
<meta charset="utf-8">
```

Aynı sayfada iki farklı charset bildirimi bırakılmamalıdır.

## Aşama 4 — Kaynak dosya dönüşümü

Bütün `.asp`, `.inc`, `.js`, `.css`, `.xml`, `.txt` dosyaları gerçek UTF-8 olarak kaydedilir.

Kontroller:

- Türkçe metinler gerçek karakterlerle yazılabilir: `Özüm Klima`, `Keşif`, `Bakım`.
- Eski HTML entity kullanımı çalışmaya devam eder fakat zorunlu değildir.
- Dosyalarda `windows-1254`, `iso-8859-9`, `CodePage=1254`, `Charset="windows-1254"` kalmamalıdır.

## Aşama 5 — Form ve veritabanı yazma testi

Staging üzerinde aşağıdaki metin admin panelinden kaydedilir ve tekrar okunur:

```text
İı Şş Ğğ Üü Öö Çç — “tırnak” €
Antalya Döşemealtı klima keşfi
```

Test zinciri:

1. Tarayıcı formu gönderir.
2. ASP `Request.Form` değeri doğru okur.
3. ADODB değeri Access'e doğru yazar.
4. Aynı kayıt panelde ve public sayfada doğru görünür.
5. Sayfa kaynak kodunda UTF-8 header ve meta aynıdır.

SQL birleştirme yerine parametreli `ADODB.Command` tercih edilmelidir. Encoding geçişinden bağımsız olarak apostrof ve injection sorunlarını azaltır.

## Aşama 6 — Kanıtlanmış bozuk kayıtların onarımı

Yalnızca audit ve görsel doğrulama ile bozuk olduğu kesinleşen kolonlar düzeltilir.

Olası eski Türkçe mapping:

- `ý` -> `ı`
- `Ý` -> `İ`
- `þ` -> `ş`
- `Þ` -> `Ş`
- `ð` -> `ğ`
- `Ð` -> `Ğ`

Bu mapping her kayıt için otomatik doğru olmayabilir. Özellikle `i/ı` ve büyük harfler bağlama göre kontrol edilmelidir.

Önce dry-run raporu, sonra kopya DB üzerinde update, en son canlı DB uygulanır.

## Aşama 7 — Public rollout

1. `charset.asp` 65001'e geçirilir.
2. Public giriş sayfaları UTF-8 directive/meta ile deploy edilir.
3. Admin paneli aynı deploy içinde geçirilir; panel 1254, public 65001 bırakılmaz.
4. IIS response header kontrol edilir.
5. Cloudflare/cache temizlenir.
6. Ana sayfa, kategori, detay, form, XML sitemap, robots ve JSON-LD test edilir.

## Kabul kriterleri

- Response header: `Content-Type: text/html; charset=utf-8`
- HTML head: yalnızca `<meta charset="utf-8">`
- Türkçe test metni kayıt/okuma turunu kayıpsız tamamlıyor
- Veritabanında yeni mojibake oluşmuyor
- Eski kayıtlar doğru görüntüleniyor
- Sitemap XML parser hatası vermiyor
- JSON-LD UTF-8 olarak parse ediliyor
- Form e-postaları Türkçe karakterleri koruyor
- 500 hatası yok

## Geri dönüş

Sorun halinde:

1. Siteyi bakım moduna alın.
2. Kod deploy'unu önceki commit'e geri alın.
3. Migration sırasında DB update yapıldıysa, migration öncesi `.mdb` yedeğini geri yükleyin.
4. Cache'i temizleyin.

Kod rollback ile DB rollback birbirinden ayrı düşünülmelidir.