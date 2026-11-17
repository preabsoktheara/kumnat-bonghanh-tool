# ប្រព័ន្ធបង្កើត «កំណត់បង្ហាញ» ស្វ័យប្រវត្តិ (AI)
**ការិយាល័យរដ្ឋបាល • សេនាធិការដ្ឋាន • ស្នងការដ្ឋាននគរបាលខេត្តព្រះសីហនុ**

## របៀបបង្ហោះទៅកាន់ Vercel (យ៉ាងងាយស្រួលបំផុត)

1. **បញ្ចូលកូដទៅ GitHub (ឬ Google AI Studio):**
   - បង្កើត Repository ថ្មីមួយនៅលើ **GitHub** (ឧទាហរណ៍៖ `kumnat-bonghanh-tool`) រួច Upload ឯកសារទាំងអស់ក្នុងកញ្ចប់នេះចូល។
2. **ភ្ជាប់ជាមួយ Vercel (`vercel.com`):**
   - ចូលទៅកាន់ [vercel.com](https://vercel.com) -> ចុច **Add New Project** -> ជ្រើសរើស Repository `kumnat-bonghanh-tool` របស់អ្នក។
3. **បន្ថែម API Key ក្នុង Vercel (សំខាន់):**
   - នៅត្រង់ផ្នែក **Environment Variables** មុនពេលចុច Deploy សូមបន្ថែម៖
     - **Name:** `VITE_GEMINI_API_KEY`
     - **Value:** បិទភ្ជាប់ Gemini API Key របស់អ្នក
4. **ចុច Deploy:**
   - រង់ចាំប្រហែល ១ នាទី អ្នកនឹងទទួលបានតំណភ្ជាប់វេបសាយ (URL) សម្រាប់ឱ្យមន្ត្រីក្នុងការិយាល័យប្រើប្រាស់រួមគ្នាបានភ្លាមៗ!
