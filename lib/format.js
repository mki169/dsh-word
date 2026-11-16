function bytes(value) {
    if (value < 1024)
        return `${value} B`;
    if (value < 1024 * 1024)
        return `${(value / 1024).toFixed(1)} KB`;
    return `${(value / (1024 * 1024)).toFixed(1)} MB`;
}
export function statsLine(stats) {
    const parts = [
        `${stats.sections} section${stats.sections === 1 ? '' : 's'}`,
        `${stats.paragraphs} paragraph${stats.paragraphs === 1 ? '' : 's'}`,
        `${stats.tables} table${stats.tables === 1 ? '' : 's'}`,
        `${stats.images} image${stats.images === 1 ? '' : 's'}`,
        `${stats.words} words`,
        bytes(stats.bytes),
    ];
    return parts.join(' · ');
}
export function formatOutline(outline, limit = 30) {
    if (outline.length === 0)
        return '';
    const lines = outline.slice(0, limit).map(item => `${'  '.repeat(Math.max(0, item.level - 1))}- [p${item.index}] ${item.text}`);
    if (outline.length > limit)
        lines.push(`- ... ${outline.length - limit} more headings`);
    return lines.join('\n');
}
export function formatRead(result) {
    const properties = result.properties;
    const meta = [properties.title && `title: ${properties.title}`, properties.author && `author: ${properties.author}`, properties.subject && `subject: ${properties.subject}`]
        .filter((value) => Boolean(value));
    const outline = formatOutline(result.outline);
    const header = result.header ? `header: ${result.header}` : '';
    const footer = result.footer ? `footer: ${result.footer}` : '';
    const lines = [
        result.path,
        statsLine(result.stats),
        ...meta,
        ...(header ? [header] : []),
        ...(footer ? [footer] : []),
        ...(outline ? ['', 'Outline:', outline] : []),
        '',
        result.markdown,
        ...(result.truncated ? ['', '[markdown truncated; pass a larger max_chars or read a narrower range]'] : []),
    ];
    return lines.join('\n');
}
export function formatCreate(result) {
    const backup = result.backup ? `\nPrevious file backed up to ${result.backup}` : '';
    return `Created ${result.path}\n${statsLine(result.stats)}${backup}`;
}
export function formatEdit(result) {
    const changes = result.changes.map(change => `- op ${change.op} ${change.action}: ${change.detail}`);
    const backup = result.backup ? `\nBackup: ${result.backup}` : '';
    return [
        `Edited ${result.path}`,
        ...changes,
        statsLine(result.stats) + backup,
    ].join('\n');
}
export function formatExport(result) {
    return `Exported PDF with ${result.engine}: ${result.output} (${bytes(result.bytes)})\nSource: ${result.input}`;
}
export function formatRender(result, total) {
    const pages = result.pages.map(page => `- page ${page.page}: ${page.width}x${page.height}`).join('\n');
    const capped = total < result.total_pages ? `\nRendered ${total} of ${result.total_pages} pages.` : '';
    return `Rendered ${result.pdf}\n${pages}${capped}`;
}
export function formatApp(result) {
    if (result.action === 'open')
        return result.message;
    const version = result.version ? ` ${result.version}` : '';
    const document = result.document;
    const open = result.openInWord === null ? '' : result.openInWord ? ' (open in Word)' : ' (not open in Word)';
    return [
        `Microsoft Word${version}`,
        `App: ${result.appPath}`,
        `Installed: ${result.installed ? 'yes' : 'no'}`,
        ...(document ? [`Document: ${document}${open}`] : []),
        result.message,
    ].join('\n');
}
//# sourceMappingURL=format.js.map