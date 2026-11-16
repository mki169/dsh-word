import { resolveConfig } from './config.js';
import { WordEngine } from './engine.js';
import { createTools } from './tools.js';
export { Config, ENTRY_ID, resolveConfig } from './config.js';
export { WordError } from './engine.js';
export const name = 'dsh-office-word';
export const inject = ['tools', 'attachments'];
export function apply(ctx, config = {}) {
    if (process.platform !== 'darwin') {
        ctx.logger.warn('dsh-office-word supports macOS only; no tools registered.');
        return;
    }
    const settings = resolveConfig(config);
    const engine = new WordEngine(settings, ctx.logger);
    const saveImage = (data, imageName) => {
        return ctx.attachments.saveImage({ data, mediaType: 'image/png', name: imageName });
    };
    for (const tool of createTools({ engine, saveImage })) {
        ctx.effect(() => ctx.tools.register(tool), `office-word: tool ${tool.name}`);
    }
}
//# sourceMappingURL=index.js.map