# Rob's Class App

A lightweight, single-file Progressive Web App (PWA) designed for educators to easily create, distribute, and grade quizzes and homework assignments. Students complete quizzes on any device via shareable links, and scores, answers, and integrity logs are automatically sent to a connected Google Sheet.

## Features

* **Spreadsheet Import:** Quickly upload quiz questions, options, answers, and points using standard Excel (`.xlsx`, `.xls`) or CSV files.
* **Shareable Quiz Links:** Encode quizzes, deadlines, and rules directly into a URL hash so students can access them instantly without needing a server backend.
* **Instant Grading & Google Sheets Sync:** Automatically grades student responses and sends results directly to a Google Sheet via a Google Apps Script webhook. Offline fallback saves results locally and retries automatically.
* **Activity & Integrity Monitoring:** Tracks and flags student activity—including page leaves, tab switching, copy/paste attempts, print shortcuts, and completion speed.
* **Customizable Deadlines & Timers:** Set strict deadlines, allow or disallow late submissions, and enforce custom time limits.
* **Student Review Mode:** Optionally allow students to view correct answers and breakdown details immediately after submission.
* **Zero Dependencies Required to Host:** Built with vanilla JavaScript, Tailwind CSS, and SheetJS, packaged entirely into a single `index.html` file.

## Quick Start for Teachers

1. **Host the App:** Upload `index.html`, `manifest.webmanifest`, and your icon files to your GitHub repository (or any static web host).
2. **Connect a Google Sheet:**
   * Open your Google Sheet, go to **Extensions > Apps Script**, and paste the provided Apps Script code.
   * Deploy it as a **Web App** (Execute as: *Me*, Who has access: *Anyone*).
   * Copy the resulting Web App URL.
3. **Configure Your Quiz:** Open the app, log in with your teacher password, upload your quiz spreadsheet, paste your webhook URL, and customize your settings.
4. **Share:** Click **Copy student link** and distribute it to your class.

## License

Open-source and free for educators.
