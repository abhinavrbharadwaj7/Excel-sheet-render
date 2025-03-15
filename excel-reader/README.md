# Excel Sheet Render

## Overview
Excel Sheet Render is a web-based tool designed to parse and display Excel sheets in a user-friendly format. This project enables users to upload Excel files and view their contents directly in the browser.

## Features
- Upload Excel files (.xlsx, .xls) for rendering
- Display sheet data in a structured table format
- Support for multiple sheets within an Excel file
- Pagination for large datasets
- User-friendly interface

## Installation

### Prerequisites
Ensure you have the following installed:
- Node.js (v14+ recommended)
- npm or yarn

### Steps
1. Clone the repository:
   ```sh
   git clone https://github.com/abhinavrbharadwaj7/Excel-sheet-render.git
   ```
2. Navigate to the project directory:
   ```sh
   cd Excel-sheet-render
   ```
3. Install dependencies:
   ```sh
   npm install  # or yarn install
   ```
4. Start the development server:
   ```sh
   npm start  # or yarn start
   ```

## Usage
1. Open the application in the browser (usually at `http://localhost:3000`).
2. Click on the **Upload** button to select an Excel file.
3. The data will be displayed in a structured table format.
4. Use the sheet selector to navigate between different sheets.

## Folder Structure
```
Excel-sheet-render/
│── public/
│── src/
│   │── components/  # Reusable UI components
│   │── pages/       # Page-level components
│   │── utils/       # Helper functions (Excel parsing, etc.)
│   │── App.tsx      # Main application file
│   │── index.tsx    # Entry point
│── package.json
│── README.md
```

## Contributing
Contributions are welcome! Feel free to fork this repository and submit a pull request with improvements.

## License
This project is licensed under the MIT License.

## Contact
For any inquiries, reach out to Abhinav R Bharadwaj at [abhinavrbharadwaj86@gmail.com].

