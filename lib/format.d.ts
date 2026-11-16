import type { AppResult, CreateResult, EditResult, ExportResult, OutlineItem, ReadResult, RenderResult, WordStats } from './types.js';
export declare function statsLine(stats: WordStats): string;
export declare function formatOutline(outline: OutlineItem[], limit?: number): string;
export declare function formatRead(result: ReadResult): string;
export declare function formatCreate(result: CreateResult): string;
export declare function formatEdit(result: EditResult): string;
export declare function formatExport(result: ExportResult): string;
export declare function formatRender(result: RenderResult, total: number): string;
export declare function formatApp(result: AppResult): string;
