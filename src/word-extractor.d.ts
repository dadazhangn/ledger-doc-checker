declare module 'word-extractor' {
  interface ExtractedWord {
    getBody(options?: { filterUnicode?: boolean }): string;
    getHeaders(options?: { filterUnicode?: boolean; includeFooters?: boolean }): string;
    getFooters(options?: { filterUnicode?: boolean }): string;
    getTextboxes(options?: { filterUnicode?: boolean; includeHeadersAndFooters?: boolean; includeBody?: boolean }): string;
  }
  export default class WordExtractor {
    extract(source: Buffer): Promise<ExtractedWord>;
  }
}
