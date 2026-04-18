import { describe, expect, it } from "vitest";
import {
  dateDetector,
  monthDetector,
  numericDetector,
  pickSeries,
  repeatDetector,
  textNumericDetector,
  weekdayDetector,
  type SeriesSample,
} from "./series.js";

function s(...values: ReadonlyArray<unknown>): SeriesSample[] {
  return values.map((v) => ({ value: v as SeriesSample["value"] }));
}

describe("numericDetector", () => {
  it("infers step from two samples", () => {
    const g = numericDetector(s(1, 2));
    expect(g).not.toBeNull();
    expect(g!.next(1)).toBe(3);
    expect(g!.next(3)).toBe(5);
  });

  it("supports non-unit steps", () => {
    const g = numericDetector(s(2, 4, 6));
    expect(g).not.toBeNull();
    expect(g!.next(1)).toBe(8);
    expect(g!.next(2)).toBe(10);
  });

  it("rejects non-arithmetic series", () => {
    expect(numericDetector(s(1, 2, 4))).toBeNull();
  });

  it("returns null for a single sample (handled by repeat)", () => {
    expect(numericDetector(s(7))).toBeNull();
  });
});

describe("repeatDetector", () => {
  it("cycles a single sample", () => {
    const g = repeatDetector(s("hi"));
    expect(g!.next(1)).toBe("hi");
    expect(g!.next(5)).toBe("hi");
  });

  it("cycles multi-sample patterns", () => {
    const g = repeatDetector(s("a", "b"));
    expect(g!.next(1)).toBe("a");
    expect(g!.next(2)).toBe("b");
    expect(g!.next(3)).toBe("a");
  });
});

describe("weekdayDetector", () => {
  it("extends short-form weekdays", () => {
    const g = weekdayDetector(s("Mon", "Tue"));
    expect(g!.next(1)).toBe("Wed");
    expect(g!.next(5)).toBe("Sun");
    expect(g!.next(6)).toBe("Mon");
  });

  it("extends long-form weekdays", () => {
    const g = weekdayDetector(s("Friday"));
    // Single sample defaults to step 1.
    expect(g!.next(1)).toBe("Saturday");
    expect(g!.next(2)).toBe("Sunday");
  });

  it("preserves casing", () => {
    expect(weekdayDetector(s("mon"))!.next(1)).toBe("tue");
    expect(weekdayDetector(s("MON"))!.next(1)).toBe("TUE");
  });

  it("rejects non-weekday strings", () => {
    expect(weekdayDetector(s("Yesterday"))).toBeNull();
  });
});

describe("monthDetector", () => {
  it("extends short months", () => {
    const g = monthDetector(s("Jan", "Feb"));
    expect(g!.next(1)).toBe("Mar");
    expect(g!.next(11)).toBe("Jan");
  });

  it("extends long months", () => {
    expect(monthDetector(s("November"))!.next(1)).toBe("December");
    expect(monthDetector(s("November"))!.next(2)).toBe("January");
  });
});

describe("textNumericDetector", () => {
  it("increments suffix integers", () => {
    const g = textNumericDetector(s("Item 1", "Item 2"));
    expect(g!.next(1)).toBe("Item 3");
  });

  it("preserves zero-padding", () => {
    const g = textNumericDetector(s("Q01", "Q02"));
    expect(g!.next(1)).toBe("Q03");
    expect(g!.next(8)).toBe("Q10");
  });

  it("rejects mismatched prefixes", () => {
    expect(textNumericDetector(s("A1", "B2"))).toBeNull();
  });
});

describe("dateDetector", () => {
  it("steps ISO dates by one day by default", () => {
    const g = dateDetector(s("2024-01-30", "2024-01-31"));
    expect(g!.next(1)).toBe("2024-02-01");
    expect(g!.next(2)).toBe("2024-02-02");
  });

  it("supports week-step dates", () => {
    const g = dateDetector(s("2024-01-01", "2024-01-08"));
    expect(g!.next(1)).toBe("2024-01-15");
    expect(g!.next(2)).toBe("2024-01-22");
  });

  it("rejects mixed steps", () => {
    expect(dateDetector(s("2024-01-01", "2024-01-02", "2024-01-04"))).toBeNull();
  });
});

describe("pickSeries priority", () => {
  it("prefers numeric over repeat for arithmetic input", () => {
    expect(pickSeries(s(1, 2, 3)).id).toBe("numeric");
  });
  it("uses weekday before text-numeric", () => {
    expect(pickSeries(s("Mon", "Tue")).id).toBe("weekday");
  });
  it("falls back to repeat for opaque mixed content", () => {
    expect(pickSeries(s("foo", 1)).id).toBe("repeat");
  });
  it("uses repeat for a single numeric sample", () => {
    expect(pickSeries(s(42)).id).toBe("repeat");
  });
});
