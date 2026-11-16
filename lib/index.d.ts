import type { Context } from '@deepseek-ai/cordis';
import { type Config as ConfigShape } from './config.js';
export { Config, ENTRY_ID, resolveConfig } from './config.js';
export { WordError } from './engine.js';
export type { ConfigShape };
export declare const name = "dsh-office-word";
export declare const inject: string[];
export declare function apply(ctx: Context, config?: ConfigShape): void;
