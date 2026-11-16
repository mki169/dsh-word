import z from '@deepseek-ai/schemastery';
export declare const ENTRY_ID = "office-word";
export declare const WORD_APP_PATH = "/Applications/Microsoft Word.app";
export declare const LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice";
export type PdfEngine = 'auto' | 'word' | 'libreoffice' | 'none';
export interface Config {
    python?: string;
    workspace?: string;
    allowOutsideWorkspace?: boolean;
    backup?: boolean;
    pdfEngine?: PdfEngine;
    sofficePath?: string;
    wordPath?: string;
    timeoutMs?: number;
    maxReadChars?: number;
    renderWidth?: number;
    maxRenderPages?: number;
}
export declare const Config: z<Config>;
export interface Settings {
    python: string;
    workspace: string;
    workspaceConfigured: boolean;
    allowOutsideWorkspace: boolean;
    backup: boolean;
    pdfEngine: PdfEngine;
    sofficePath: string;
    wordPath: string;
    timeoutMs: number;
    maxReadChars: number;
    renderWidth: number;
    maxRenderPages: number;
}
export declare function resolveConfig(raw?: Config): Settings;
