import type { CommandHandler } from "./commands/types.js";
import type { DocumentFormat, DocumentSnapshot } from "./types/document.js";

export interface FormatPlugin<TSnapshot extends DocumentSnapshot = DocumentSnapshot> {
  readonly format: DocumentFormat;
  readonly handlers: ReadonlyArray<CommandHandler<unknown, TSnapshot>>;
  parse(buffer: ArrayBuffer | Uint8Array): Promise<TSnapshot>;
  serialize(snapshot: TSnapshot): Promise<ArrayBuffer>;
  toMarkdown?(snapshot: TSnapshot): string;
}

export class PluginRegistry {
  private readonly plugins = new Map<DocumentFormat, FormatPlugin<DocumentSnapshot>>();

  register<T extends DocumentSnapshot>(plugin: FormatPlugin<T>): void {
    this.plugins.set(plugin.format, plugin as unknown as FormatPlugin<DocumentSnapshot>);
  }

  get(format: DocumentFormat): FormatPlugin<DocumentSnapshot> {
    const p = this.plugins.get(format);
    if (!p) throw new Error(`No plugin registered for format "${format}"`);
    return p;
  }

  has(format: DocumentFormat): boolean {
    return this.plugins.has(format);
  }
}
