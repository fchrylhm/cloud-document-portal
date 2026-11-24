# Cloud Document Portal (Auto-Organizer Ormawa & Cloud Verification)

Aplikasi web *serverless* berbasis ekosistem Google Cloud untuk mengelola pengumpulan, pengarsipan otomatis, dan verifikasi berkas administrasi Organisasi Mahasiswa (Ormawa) seperti Proposal, LPJ, Surat Izin, dan Laporan Keuangan.

Proyek ini dibangun untuk memenuhi tugas Ujian Akhir Semester (UAS) mata kuliah **Cloud Computing** dengan memanfaatkan integrasi penuh layanan cloud Google tanpa manajemen server fisik (*zero maintenance*).

---

## Arsitektur & Penerapan Konsep Cloud Computing

Aplikasi ini mengorkestrasi empat layanan *fully managed* Google dalam satu alur *serverless*:

```text
[ Browser: User / Admin (HTML + Tailwind CSS + JS) ]
                        │
                        ▼  (google.script.run - Asynchronous RPC)
[ Google Apps Script (Code.gs) - Serverless Compute ]
        ├──► Auth & Security  (SHA-256 Hashing + 2-Layer File Whitelist)
        ├──► Google Drive     (Cloud Object Storage: Auto-Folder, Auto-Rename, Lifecycle)
        ├──► Google Sheets    (Cloud Database: Users, Documents & Audit Logs)
        └──► Gmail Service    (Cloud Notification & Daily Quota Monitoring)