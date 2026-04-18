import type { NodeId } from "@officeai/core";

// ─── Common ───────────────────────────────────────────────────────────────

export interface PptxTextRange {
  readonly paragraph: number;
  readonly start: number;
  readonly end: number;
}

export interface TextFormatPayload {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string;
  readonly strike?: boolean;
  readonly fontFamily?: string;
  readonly fontSizeHundredths?: number;
  readonly color?: string;
}

// ─── P0 payloads ──────────────────────────────────────────────────────────

export interface AddSlidePayload {
  readonly at?: number;
  readonly layoutPartPath?: string;
}

export interface DeleteSlidePayload {
  readonly slideIndex: number;
}

export interface DuplicateSlidePayload {
  readonly slideIndex: number;
}

export interface MoveSlidePayload {
  readonly from: number;
  readonly to: number;
}

export interface SetTextPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly text: string;
}

export interface SetPositionPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly x: number;
  readonly y: number;
}

export interface SetSizePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly width: number;
  readonly height: number;
}

// ─── P1 payloads ──────────────────────────────────────────────────────────

export interface FormatTextPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly range: PptxTextRange;
  readonly format: TextFormatPayload;
}

export interface InsertImagePayload {
  readonly slideIndex: number;
  readonly data: Uint8Array | ArrayBuffer;
  readonly mimeType: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly altText?: string;
  readonly name?: string;
}

export interface AddTextBoxPayload {
  readonly slideIndex: number;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly name?: string;
}

// ─── F2 (Tables) payloads ─────────────────────────────────────────────────

export interface TableSetCellTextPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly row: number;
  readonly column: number;
  readonly text: string;
}

export interface TableAddRowPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly at?: number;
  readonly height?: number;
}

export interface TableDeleteRowPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly row: number;
}

export interface TableAddColumnPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly at?: number;
  readonly width?: number;
}

export interface TableDeleteColumnPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly column: number;
}

// ─── F3 (Charts) payloads ─────────────────────────────────────────────────

export interface SetChartTitlePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  /** New title; pass `null` to remove the title. */
  readonly title: string | null;
}

export interface SetChartDataPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<{
    readonly name?: string;
    readonly values: ReadonlyArray<number>;
  }>;
}

export interface SetChartTypePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly chartType: "bar" | "line" | "pie" | "area";
}

// ─── Type tags ────────────────────────────────────────────────────────────

export const PPTX_COMMAND_TYPES = {
  addSlide: "pptx:add-slide",
  deleteSlide: "pptx:delete-slide",
  duplicateSlide: "pptx:duplicate-slide",
  moveSlide: "pptx:move-slide",
  setText: "pptx:set-text",
  setPosition: "pptx:set-position",
  setSize: "pptx:set-size",
  formatText: "pptx:format-text",
  insertImage: "pptx:insert-image",
  addTextBox: "pptx:add-text-box",
  tableSetCellText: "pptx:table-set-cell-text",
  tableAddRow: "pptx:table-add-row",
  tableDeleteRow: "pptx:table-delete-row",
  tableAddColumn: "pptx:table-add-column",
  tableDeleteColumn: "pptx:table-delete-column",
  setChartTitle: "pptx:set-chart-title",
  setChartData: "pptx:set-chart-data",
  setChartType: "pptx:set-chart-type",
} as const;

export type PptxCommandType = (typeof PPTX_COMMAND_TYPES)[keyof typeof PPTX_COMMAND_TYPES];
