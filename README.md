# Sistem OPR PLC — Sekolah Kebangsaan Seri Setia

Pakej ini sedia diterbitkan melalui GitHub Pages.

## Cara menerbitkan

1. Muat naik `index.html` dan folder `assets` ke bahagian utama repository GitHub.
2. Buka **Settings → Pages**.
3. Pada **Build and deployment**, pilih **Deploy from a branch**.
4. Pilih branch `main` dan folder `/(root)`, kemudian klik **Save**.
5. Tunggu sehingga pautan GitHub Pages dipaparkan.

Jangan ubah nama atau lokasi folder `assets` kerana header dan logo menggunakan fail di dalam folder tersebut.

## Status sambungan

- CSV rekod OPR telah dikonfigurasikan kepada `gid=0`.
- CSV nama guru telah dikonfigurasikan kepada `gid=1130301879`.
- URL Apps Script Web App baharu telah dimasukkan ke dalam `index.html`.
- Pastikan `Code.gs` menggunakan Sheet ID `11F7GYjSjHcrVafEet_LLrzHqytpQ9f5lEI13P4D1XvM` dan Drive ID `1oCNCL11GL1ane_LQDmN_SshSUc3EWF3n`.

## Apps Script

Fail `Code.gs` yang lengkap turut disertakan. Jika perlu menggantikan backend, salin seluruh kandungannya ke projek Apps Script, jalankan `setupSistem`, kemudian deploy sebagai Web app (`Execute as: Me`, `Who has access: Anyone`).
