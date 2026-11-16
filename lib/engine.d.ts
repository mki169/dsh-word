import { HarnessError } from '@deepseek-ai/dsh-llm';
import type { PdfEngine, Settings } from './config.js';
import type { AppResult, CreateResult, EditResult, ExportResult, ProbeResult, ReadResult, RenderResult } from './types.js';
export declare class WordError extends HarnessError {
    constructor(code: string, message: string, options?: ErrorOptions);
}
export interface LoggerLike {
    warn(message: string): void;
}
interface ProcessOptions {
    input?: string;
    timeoutMs?: number;
    signal?: AbortSignal;
    env?: NodeJS.ProcessEnv;
    cwd?: string;
}
interface ProcessResult {
    stdout: string;
    stderr: string;
    code: number | null;
}
export declare function runProcess(command: string, args: string[], options?: ProcessOptions): Promise<ProcessResult>;
export type PathMode = 'read' | 'write' | 'create' | 'pdf';
export declare class WordEngine {
    readonly settings: Settings;
    private readonly logger;
    private readonly helperPath;
    private readonly wordScriptPath;
    private probeCache?;
    private readonly locks;
    constructor(settings: Settings, logger: LoggerLike);
    resolvePath(input: string, mode: PathMode, workspace?: string): Promise<string>;
    resolveAssetPath(input: string, workspace?: string): Promise<string>;
    private assertAllowed;
    probe(signal?: AbortSignal): Promise<ProbeResult>;
    ensureReady(signal?: AbortSignal): Promise<ProbeResult>;
    ensureRenderReady(signal?: AbortSignal): Promise<ProbeResult>;
    read(path: string, options?: {
        includeHeaderFooter?: boolean;
        maxChars?: number;
    }, signal?: AbortSignal): Promise<ReadResult>;
    create(path: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<CreateResult>;
    edit(path: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<EditResult>;
    private stagingDirectory;
    exportPdf(input: string, output: string, engine: PdfEngine, overwrite: boolean, signal?: AbortSignal): Promise<ExportResult>;
    renderPages(document: string, options?: {
        pages?: string;
        width?: number;
        engine?: PdfEngine;
        maxPages?: number;
    }, signal?: AbortSignal): Promise<RenderResult & {
        cleanup: () => Promise<void>;
    }>;
    appStatus(document?: string, signal?: AbortSignal): Promise<AppResult>;
    openInWord(document: string, signal?: AbortSignal): Promise<AppResult>;
    isOpenInWord(document: string, signal?: AbortSignal): Promise<boolean | null>;
    private resolveEngine;
    private runHelper;
    private withLock;
}
export {};
