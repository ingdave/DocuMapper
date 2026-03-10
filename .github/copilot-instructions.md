# Document Generator - Copilot Instructions

This workspace contains a full-stack web application for mapping Excel data to Word templates.

## Project Structure

- **backend/**: Express.js + TypeScript API server
- **frontend/**: Angular 16 web application

## Development Commands

### Backend
```bash
cd backend
npm install
npm run dev      # Start with ts-node
npm run build    # Compile TypeScript
npm run start    # Run compiled version
```

### Frontend
```bash
cd frontend
npm install
npm start        # Dev server on http://localhost:4200
npm run build    # Production build
```

## Key Technologies

- **Backend**: Node.js, Express, TypeScript, exceljs, docxtemplater
- **Frontend**: Angular 16, TypeScript, SCSS, RxJS
- **Protocol**: REST API on port 3000

## Important Services

- **ApiService** (`frontend/src/app/services/api.service.ts`) - HTTP calls to backend
- **StateService** (`frontend/src/app/services/state.service.ts`) - Shared state management
- **Excel Route** (`backend/src/routes/excel.ts`) - Parse Excel files
- **Word Route** (`backend/src/routes/word.ts`) - Extract Word variables
- **Mapping Route** (`backend/src/routes/mapping.ts`) - Generate final documents

## Main Features to Implement

1. **Excel Processing**: Use exceljs to properly parse Excel files
2. **Word Variable Extraction**: Use docxtemplater to find placeholders
3. **Drag-Drop Mapping**: Already implemented in MappingGridComponent
4. **Document Generation**: Complete the generate endpoint

## Common Tasks

- To add a new component: Create file in `frontend/src/app/components/`
- To add a new API endpoint: Create route file in `backend/src/routes/`
- To modify styling: Edit SCSS in component files or `frontend/src/styles.scss`

## Current Status

All scaffold files created. Next: Install dependencies and test compilation.
