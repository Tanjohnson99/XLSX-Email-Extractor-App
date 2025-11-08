# DataPilot Workspace

DataPilot is a browser-based productivity workspace designed for handling spreadsheet-driven workflows. The first module in the suite is the **Email Extractor**, a tool that parses XLSX files, aggregates valid addresses, and provides instant export-ready lists.

## Table of Contents
- [Features](#features)
- [Live Demo](#live-demo)
- [Quick Start](#quick-start)
- [Using the Email Extractor](#using-the-email-extractor)
- [Interface Overview](#interface-overview)
- [Development](#development)
- [Project Structure](#project-structure)
- [Contributing](#contributing)
- [License](#license)

## Features
- Modern SaaS-inspired interface with persistent sidebar navigation and responsive cards.
- XLSX email extraction with deduplication, validation, and export to plain-text files.
- Intelligent status messaging that guides users through file selection, processing, and download.
- Mobile-friendly layout that swaps the sidebar for a hamburger menu and retains full functionality on small screens.

## Live Demo
Explore the workspace experience in the animated preview:

![Email extractor demo](demo.gif)

## Quick Start
1. **Clone** the repository
   ```bash
   git clone https://github.com/alpozkanm/XLSX-Email-Extractor-App.git
   cd XLSX-Email-Extractor-App
   ```
2. **Install dependencies**
   No build step is required—the project runs entirely in the browser.
3. **Launch** the app
   Open `index.html` in your preferred web browser.

> **Tip:** Serve the project with a local web server (such as `npx serve` or VS Code Live Server) for the best experience, especially when testing on mobile devices.

## Using the Email Extractor
1. Click **Upload XLSX** and choose one or multiple Excel files.
2. The extractor parses every sheet, finds unique email addresses, and displays the results in the **Email Extractor** card.
3. Use the **Copy** button for quick clipboard access or **Download List** to export a `.txt` file containing the collected addresses.
4. Need to start over? Use **Clear Results** to reset the state and upload new files.

## Interface Overview
- **Sidebar / Hamburger Menu** – Navigate between workspace modules (additional slots are ready for future tools).
- **Metric Panels** – Highlight the number of processed files, total email count, and time of the last extraction.
- **Activity Feed** – Review a log of recent uploads and extraction outcomes.
- **Content Cards** – Each module appears as a dedicated card with actions, descriptions, and contextual help.

The layout adapts seamlessly to mobile screens. The sidebar collapses into a hamburger menu, keeping primary actions accessible in one tap.

## Development
- **Technologies:** Vanilla HTML, CSS, and JavaScript with the bundled [`xlsx.full.min.js`](./xlsx.full.min.js) library.
- **Code style:** The UI leverages CSS custom properties and utility classes to keep theming consistent across components.
- **Extensibility:** Additional modules can be introduced by creating new cards and linking them through the sidebar navigation script found in [`script.js`](./script.js).

### Local testing
Run a lightweight static file server to test on multiple devices:
```bash
npx serve .
```
Then open the printed URL in your browser (or mobile device on the same network).

## Project Structure
```
├── index.html          # Workspace layout and module containers
├── styles.css          # Theming, responsive layout, and component styles
├── script.js           # Navigation, menu toggles, and email extraction logic
├── xlsx.full.min.js    # XLSX parsing library
├── demo.gif            # Animated demo of the extractor
└── demo-assets/        # Source assets for the animated demo
```

## Contributing
Contributions are welcome! Feel free to open issues for bugs, feature requests, or UX improvements, and submit pull requests with proposed changes.

## License
Released under the [MIT License](LICENSE).
