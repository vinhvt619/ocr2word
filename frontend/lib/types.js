/**
 * @typedef {"uploading" | "queued" | "processing" | "retrying" | "packaging" | "complete" | "error"} JobStatus
 */

/**
 * @typedef {Object} JobFile
 * @property {string} name
 * @property {number} size
 * @property {string} content_type
 * @property {string} pathname
 */

/**
 * @typedef {Object} JobOptions
 * @property {"docx" | "xlsx"} [output_format]
 * @property {string} language
 * @property {"auto" | "force" | "text"} ocr_mode
 * @property {number} dpi
 * @property {boolean} page_breaks
 * @property {boolean} preserve_images
 * @property {boolean} preserve_layout
 * @property {boolean} detect_tables
 * @property {boolean} [math_ocr]
 * @property {"rules" | "local_hybrid"} [math_engine]
 */

/**
 * @typedef {Object} JobRecord
 * @property {string} job_id
 * @property {JobStatus} status
 * @property {number} percent
 * @property {string} message
 * @property {string | null} current_file
 * @property {number} completed_files
 * @property {number} total_files
 * @property {string} created_at
 * @property {string} updated_at
 * @property {string} expires_at
 * @property {JobFile[]} files
 * @property {JobOptions} options
 * @property {string} output_pathname
 * @property {string} output_name
 * @property {string} output_content_type
 * @property {Array<Record<string, unknown>>} report
 * @property {string | null} error
 * @property {"legacy" | "analyzing_pdf" | "chunked_pdf"} [pipeline]
 * @property {number} [source_page_count]
 * @property {number} [ocr_page_count]
 * @property {number} [chunk_total]
 * @property {number} [chunk_completed]
 */

/**
 * @typedef {Object} ConvertQueueMessage
 * @property {string} job_id
 * @property {"analyze" | "convert" | "chunk" | "merge"} [action]
 * @property {number} [chunk_index]
 */

/**
 * @typedef {Object} PdfChunkRecord
 * @property {string} job_id
 * @property {number} index
 * @property {number} start_page
 * @property {number} end_page
 * @property {string} pathname
 * @property {"queued" | "processing" | "complete" | "error"} status
 * @property {number} size
 * @property {Array<Record<string, unknown>>} report
 * @property {string | null} error
 */

/**
 * @typedef {Object} CleanupRecord
 * @property {number} expires_at_ms
 * @property {string[]} pathnames
 */

export {};
