# Excel Barcode Modul

## Bikin Barcode Langsung di Excel? Bisa Banget!

Modul VBA ini lahir dari kebutuhan nyata di lapangan, banyak dari user Excel yang berkeinginan untuk bisa membuat barcode langsung dari worksheet, tanpa harus buka software lain atau ribet install AddIn tambahan.

## Solusinya?

Modul ini memberikan kamu fungsi-fungsi khusus yang bisa dipanggil langsung dari cell pada formula Excel, seperti memanggil rumus biasa. Praktis sekali untuk membuat label, mengatur stok, mencetak barcode buat produk, atau sekadar untuk testing format barcode yang beda-beda.

## Barcode yang Didukung

Pada modul ini, Anda dapat men-generate berbagai jenis barcode, tinggal dipilih sesuai kebutuhan:

* ✅ EAN8, EAN13
* ✅ UPCA, UPCE
* ✅ Code39, Code128
* ✅ 2of5 Interleaved
* ✅ DataMatrix
* ✅ QR Code (dengan pilihan Error Correction Level: L, M, Q, H)

## Cara Pakainya

Gampang banget. Di cell Excel, tinggal panggil fungsi sesuai jenis barcode yang kamu mau. Contoh:

```vba
=GenerateEAN8("1234567")
```

Atau kalau mau ambil teks dari cell lain, seperti cell "A1" :

```vba
=GenerateEAN8(A1)
```

Semua fungsi bisa dipakai langsung seperti rumus Excel biasa, cukup pastikan bahwa modulnya sudah di-load.

## Format Lengkap

Berikut daftar fungsi yang bisa kamu pakai:

```vba
=GenerateEAN13("123456789012")
=GenerateUPCA("12345678901")
=GenerateUPCE("1234567")
=Generate2of5I("123456")
=GenerateCode39("ABC123")
=GenerateCode128("ABC123")
=GenerateDataMatrix("ABC123")
=GenerateQRCode_L("ABC123")
=GenerateQRCode_M("ABC123")
=GenerateQRCode_Q("ABC123")
=GenerateQRCode_H("ABC123")
```

Sebelum Anda dapat menggunakan formula tersebut, terlebih dahulu Anda harus mengimport modul nya terlebih dahulu dengan cara:

## Langkah 1: Unduh dan Ekstrak Modul Barcode

1. Unduh modul barcode dari tautan yang tersedia.
2. Ekstrak file ke komputer Anda.

## Langkah 2: Aktifkan Tab Developer di Excel

1. Buka Microsoft Excel dan pilih **Blank Workbook** .
2. Klik **File** > **Options** > **Customize Ribbon** .
3. Di bagian **Main Tabs** , centang opsi **Developer** , lalu klik **OK** .
4. Tab **Developer** kini akan muncul di Excel Anda.

## Langkah 3: Impor Modul Barcode ke Editor VBA

1. Klik **Developer** > **View Code** atau tekan **ALT + F11** untuk membuka Editor VBA.
2. Di jendela Editor, klik kanan pada **VBAProject** , pilih **Import File** , lalu pilih file modul barcode (**modBarcode.bas**) yang telah diunduh.

## Langkah 4: Gunakan Formula Barcode di Spreadsheet

1. Kembali ke tampilan spreadsheet Excel.
2. Tentukan teks yang ingin Anda ubah menjadi barcode.
3. Gunakan fungsi khusus (User Defined Function) yang tersedia untuk membuat barcode.

Sebagai contoh, untuk membuat barcode EAN8:

```
=GenerateEAN8("BARCODE_TEXT")
```

Ganti `BARCODE_TEXT` dengan teks yang ingin Anda ubah menjadi barcode, atau gunakan referensi sel seperti:

```
=GenerateEAN8(A1)
```

Ubah nilai pada sel A1 untuk menghasilkan barcode baru.

🎯 Modul ini sangat cocok buat kamu yang menginginkan solusi cepat, ringan, dan langsung dapat berjalan di Excel. Tidak perlu coding yang ribet, tinggal panggil formula dan barcode langsung muncul.

🚀 Yuk cobain sekarang dan bikin workflow kamu makin efisien.

[Video Tutorial](https://youtu.be/jvUccXQ46Rkhttps:/)

[Download Project](https://himaapp.myr.id/catalog/excel-barcode-modul)
