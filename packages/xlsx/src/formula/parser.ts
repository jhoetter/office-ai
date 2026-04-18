import {
  collectDependencies,
  containsVolatile,
  FormulaParseError,
  type AstNode,
  type BinaryOp,
  type Formula,
} from "./ast.js";
import type { CellError } from "./errors.js";
import { lex } from "./lexer.js";
import { parseA1, parseA1Range, type CellRef, type RangeRef } from "./references.js";
import type { Token } from "./tokens.js";
import { bool as boolValue, err, num, str } from "./values.js";

export interface ParseOptions {
  readonly anchor: CellRef;
  readonly definedNames?: ReadonlyMap<string, RangeRef | CellRef>;
}

/**
 * Parse a formula string (with or without a leading `=`) into a typed
 * `Formula`. Uses recursive-descent + precedence climbing per the
 * operator table in spec/xlsx/formula-engine.md §5.
 */
export function parse(text: string, opts: ParseOptions): Formula {
  if (text === "" || text === "=") {
    throw new FormulaParseError("empty-formula", "Formula is empty", { start: 0, end: text.length });
  }
  const tokens = lex(text);
  const p = new Parser(tokens, opts, text);
  const ast = p.parseExpression(0);
  const trailing = p.peek();
  if (trailing.type === "SEMICOLON") {
    throw new FormulaParseError(
      "intersection-operator-not-supported",
      "Range intersection (`;`) is not supported in P0; use INDEX or named ranges.",
      { start: trailing.start, end: trailing.end }
    );
  }
  p.expect("EOF");
  const dependencies = collectDependencies(ast);
  const volatile = containsVolatile(ast);
  return { text, ast, anchor: opts.anchor, dependencies, volatile };
}

interface OpInfo {
  prec: number; //  HIGHER binds tighter (standard precedence-climbing convention)
  assoc: "left" | "right";
  arity: "binary";
}

/**
 * Higher number = tighter binding. Layout (highest → lowest):
 *   8 = `^`            right-assoc
 *   7 = `*` `/`        left
 *   6 = `+` `-`        left
 *   5 = `&`            left
 *   4 = comparisons    left
 *
 * Unary `-` is handled in `parseUnary` and binds tighter than `^`
 * (Excel quirk: `-2^2 == 4`).
 */
const BINARY_PRECEDENCE: Record<BinaryOp, OpInfo> = {
  "^": { prec: 8, assoc: "right", arity: "binary" },
  "*": { prec: 7, assoc: "left", arity: "binary" },
  "/": { prec: 7, assoc: "left", arity: "binary" },
  "+": { prec: 6, assoc: "left", arity: "binary" },
  "-": { prec: 6, assoc: "left", arity: "binary" },
  "&": { prec: 5, assoc: "left", arity: "binary" },
  "=": { prec: 4, assoc: "left", arity: "binary" },
  "<>": { prec: 4, assoc: "left", arity: "binary" },
  "<": { prec: 4, assoc: "left", arity: "binary" },
  ">": { prec: 4, assoc: "left", arity: "binary" },
  "<=": { prec: 4, assoc: "left", arity: "binary" },
  ">=": { prec: 4, assoc: "left", arity: "binary" },
};

class Parser {
  private pos = 0;
  constructor(
    private readonly tokens: ReadonlyArray<Token>,
    private readonly opts: ParseOptions,
    private readonly source: string
  ) {}

  parseExpression(minPrec: number): AstNode {
    let left = this.parseUnary();
    // Handle range operator (`:`) and percent (`%`) which are not in the binary table.
    while (true) {
      const t = this.peek();
      if (t.type === "PERCENT") {
        this.advance();
        left = { kind: "pct", operand: left, start: left.start, end: t.end };
        continue;
      }
      if (t.type === "OPERATOR") {
        const op = t.value as BinaryOp;
        const info = BINARY_PRECEDENCE[op];
        if (!info) break;
        if (info.prec < minPrec) break;
        this.advance();
        const nextMin = info.assoc === "left" ? info.prec + 1 : info.prec;
        const right = this.parseExpression(nextMin);
        left = { kind: "binary", op, left, right, start: left.start, end: right.end };
        continue;
      }
      break;
    }
    return left;
  }

  parseUnary(): AstNode {
    const t = this.peek();
    if (t.type === "OPERATOR" && (t.value === "-" || t.value === "+")) {
      this.advance();
      // Unary `-` binds tighter than `^` (Excel quirk: -2^2 == 4):
      // parse the operand at a precedence level *above* `^` (which is 8).
      const operand = this.parseExpression(9);
      if (t.value === "+") return operand;
      return { kind: "unary", op: "-", operand, start: t.start, end: operand.end };
    }
    return this.parsePrimary();
  }

  parsePrimary(): AstNode {
    const t = this.peek();
    switch (t.type) {
      case "NUMBER":
        this.advance();
        return { kind: "lit", value: num(t.value as number), start: t.start, end: t.end };
      case "STRING":
        this.advance();
        return { kind: "lit", value: str(t.value as string), start: t.start, end: t.end };
      case "BOOL":
        this.advance();
        return { kind: "lit", value: boolValue(t.value as boolean), start: t.start, end: t.end };
      case "ERROR":
        this.advance();
        return {
          kind: "lit",
          value: err(t.value as CellError),
          start: t.start,
          end: t.end,
        };
      case "REF": {
        this.advance();
        const ref = parseA1(t.text, this.opts.anchor.sheet);
        if (!ref) {
          throw new FormulaParseError("invalid-ref", `Invalid cell reference "${t.text}"`, {
            start: t.start,
            end: t.end,
          });
        }
        return { kind: "ref", ref, start: t.start, end: t.end };
      }
      case "RANGE_REF": {
        this.advance();
        const ref = parseA1Range(t.text, this.opts.anchor.sheet);
        if (!ref) {
          throw new FormulaParseError("invalid-ref", `Invalid range reference "${t.text}"`, {
            start: t.start,
            end: t.end,
          });
        }
        return { kind: "range", ref, start: t.start, end: t.end };
      }
      case "NAME": {
        this.advance();
        const name = t.text;
        const resolved = this.opts.definedNames?.get(name) ?? this.opts.definedNames?.get(name.toUpperCase());
        if (resolved) {
          if ("row" in resolved) {
            return { kind: "ref", ref: resolved, start: t.start, end: t.end };
          }
          return { kind: "range", ref: resolved, start: t.start, end: t.end };
        }
        return { kind: "name", name, start: t.start, end: t.end };
      }
      case "FUNCTION": {
        this.advance();
        this.expect("LPAREN");
        const args: AstNode[] = [];
        if (this.peek().type !== "RPAREN") {
          args.push(this.parseExpression(1));
          while (this.peek().type === "COMMA") {
            this.advance();
            args.push(this.parseExpression(1));
          }
        }
        const rp = this.expect("RPAREN");
        return {
          kind: "call",
          name: (t.value as string).toUpperCase(),
          args,
          start: t.start,
          end: rp.end,
        };
      }
      case "LPAREN": {
        this.advance();
        const inner = this.parseExpression(1);
        this.expect("RPAREN");
        return inner;
      }
      case "LBRACE": {
        return this.parseArrayLiteral();
      }
      case "SEMICOLON":
        throw new FormulaParseError(
          "intersection-operator-not-supported",
          "Range intersection (`;`) is not supported in P0; use INDEX or named ranges.",
          { start: t.start, end: t.end }
        );
      case "EOF":
        throw new FormulaParseError("unexpected-eof", "Unexpected end of formula", {
          start: t.start,
          end: t.end,
        });
      case "RPAREN":
      case "RBRACE":
      case "COMMA":
      case "COLON":
      case "PERCENT":
      case "OPERATOR":
        throw new FormulaParseError("unexpected-token", `Unexpected token "${t.text}" at ${t.start}`, {
          start: t.start,
          end: t.end,
        });
    }
  }

  parseArrayLiteral(): AstNode {
    const open = this.expect("LBRACE");
    const rows: AstNode[][] = [];
    let current: AstNode[] = [];
    if (this.peek().type !== "RBRACE") {
      current.push(this.parseExpression(1));
      while (true) {
        const t = this.peek();
        if (t.type === "COMMA") {
          this.advance();
          current.push(this.parseExpression(1));
          continue;
        }
        if (t.type === "SEMICOLON") {
          this.advance();
          rows.push(current);
          current = [this.parseExpression(1)];
          continue;
        }
        break;
      }
    }
    rows.push(current);
    const close = this.expect("RBRACE");
    return { kind: "array", rows, start: open.start, end: close.end };
  }

  peek(offset = 0): Token {
    return this.tokens[this.pos + offset];
  }

  advance(): Token {
    return this.tokens[this.pos++];
  }

  expect(type: Token["type"]): Token {
    const t = this.tokens[this.pos];
    if (t.type !== type) {
      throw new FormulaParseError(
        type === "RPAREN" ? "unbalanced-paren" : "unexpected-token",
        `Expected ${type}, got ${t.type} ("${t.text}") in "${this.source}"`,
        { start: t.start, end: t.end }
      );
    }
    this.pos++;
    return t;
  }
}
