export interface WordStats {
    paragraphs: number;
    tables: number;
    images: number;
    sections: number;
    words: number;
    pages_estimate: number;
    bytes: number;
}
export interface DocProperties {
    title: string;
    author: string;
    subject: string;
    keywords: string;
    comments: string;
    last_modified_by: string;
    revision: number;
    created: string;
    modified: string;
}
export interface OutlineItem {
    index: number;
    level: number;
    style: string;
    text: string;
}
export interface ReadResult {
    path: string;
    stats: WordStats;
    properties: DocProperties;
    markdown: string;
    outline: OutlineItem[];
    truncated: boolean;
    header: string;
    footer: string;
}
export interface CreateResult {
    path: string;
    bytes: number;
    backup: string | null;
    stats: WordStats;
}
export interface EditChange {
    op: number;
    action: string;
    status: string;
    detail: string;
}
export interface EditResult {
    path: string;
    backup: string | null;
    changes: EditChange[];
    stats: WordStats;
}
export interface RenderPageInfo {
    page: number;
    path: string;
    width: number;
    height: number;
    mime: string;
}
export interface RenderResult {
    pdf: string;
    total_pages: number;
    pages: RenderPageInfo[];
}
export interface ExportResult {
    input: string;
    output: string;
    engine: 'libreoffice' | 'word';
    bytes: number;
}
export interface ProbeResult {
    helper: string;
    python: string;
    python_docx: string | null;
    pymupdf: string | null;
    platform: string;
    word_app: boolean;
    word_version: string | null;
    libreoffice: boolean;
}
export interface AppResult {
    installed: boolean;
    version: string | null;
    appPath: string;
    document: string | null;
    openInWord: boolean | null;
    action: 'status' | 'open';
    message: string;
}
export interface AttachmentJson {
    attachmentId: string;
    mediaType: string;
    bytes: number;
    width: number;
    height: number;
    name?: string;
    originalDimensions?: {
        width: number;
        height: number;
    };
}
export interface RenderValue {
    text: string;
    image?: AttachmentJson;
    more?: AttachmentJson[];
}
export interface HelperErrorBody {
    code: string;
    message: string;
}
export interface HelperFailure {
    ok: false;
    error: HelperErrorBody;
}
export interface HelperSuccess<T> {
    ok: true;
    result: T;
}
export type HelperResponse<T> = HelperSuccess<T> | HelperFailure;
