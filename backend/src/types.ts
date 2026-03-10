export interface ExcelData {
  columns: string[];
  rows: Record<string, any>[];
}

export interface WordVariable {
  name: string;
  placeholder: string;
}

export interface MappingConfig {
  mappings: Record<string, string>; // { excelColumn: wordPlaceholder }
  timestamp: number;
}

export interface GenerateDocumentRequest {
  excelData: ExcelData;
  mappingConfig: MappingConfig;
  wordTemplateBase64: string;
}
