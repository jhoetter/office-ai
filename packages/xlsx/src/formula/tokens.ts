import type { CellError } from "./errors.js";

/**
 * Token types and `Token` shape produced by the lexer.
 *
 * Spec: `spec/xlsx/formula-engine.md` §4.
 */
export type TokenType =
  | "NUMBER"
  | "STRING"
  | "BOOL"
  | "ERROR"
  | "REF"
  | "RANGE_REF"
  | "NAME"
  | "FUNCTION"
  | "OPERATOR"
  | "LPAREN"
  | "RPAREN"
  | "LBRACE"
  | "RBRACE"
  | "COMMA"
  | "COLON"
  | "SEMICOLON"
  | "PERCENT"
  | "EOF";

export interface Token {
  readonly type: TokenType;
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly value?: number | string | boolean | CellError;
  /** Set by shunting-yard for FUNCTION tokens to record arg count. */
  argCount?: number;
}
