export type JobStatus =
  | "uploading"
  | "queued"
  | "processing"
  | "retrying"
  | "packaging"
  | "complete"
  | "error";

export interface JobFile {
  name: string;
  size: number;
  content_type: string;
  pathname: string;
}

export interface JobOptions {
  language: string;
  ocr_mode: "auto" | "force" | "text";
  dpi: number;
  page_breaks: boolean;
  preserve_images: boolean;
  preserve_layout: boolean;
  detect_tables: boolean;
  ai_assist: boolean;
}

export interface JobRecord {
  job_id: string;
  status: JobStatus;
  percent: number;
  message: string;
  current_file: string | null;
  completed_files: number;
  total_files: number;
  created_at: string;
  updated_at: string;
  expires_at: string;
  files: JobFile[];
  options: JobOptions;
  output_pathname: string;
  output_name: string;
  output_content_type: string;
  report: Array<Record<string, unknown>>;
  error: string | null;
  pipeline?: "legacy" | "analyzing_pdf" | "chunked_pdf";
  source_page_count?: number;
  ocr_page_count?: number;
  chunk_total?: number;
  chunk_completed?: number;
  ai_input_pathname?: string | null;
  ai_input_size?: number;
}

export interface ConvertQueueMessage {
  job_id: string;
  action?: "analyze" | "convert" | "chunk" | "merge" | "ai";
  chunk_index?: number;
}

export interface PdfChunkRecord {
  job_id: string;
  index: number;
  start_page: number;
  end_page: number;
  pathname: string;
  status: "queued" | "processing" | "complete" | "error";
  size: number;
  report: Array<Record<string, unknown>>;
  error: string | null;
}

export interface CleanupRecord {
  expires_at_ms: number;
  pathnames: string[];
}
