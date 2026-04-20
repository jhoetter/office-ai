// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";

// React 19 wants this flag so `act(...)` warnings don't fire on
// every render. Set BEFORE the first `createRoot` call.
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;
import {
  ResizeHandles,
  isCornerSide,
  type ResizeHandleGrabInfo,
  type ResizeHandleSide,
} from "./resize-handles";

let host: HTMLDivElement;
let root: Root;

beforeEach(() => {
  host = document.createElement("div");
  // The primitive expects a positioned ancestor; mirror what real
  // consumers do.
  host.style.position = "relative";
  document.body.appendChild(host);
  root = createRoot(host);
});

afterEach(() => {
  act(() => {
    root.unmount();
  });
  host.remove();
});

function render(ui: React.ReactElement): void {
  act(() => {
    root.render(ui);
  });
}

function fireMouseDown(el: Element, opts: MouseEventInit = {}): void {
  act(() => {
    el.dispatchEvent(
      new MouseEvent("mousedown", {
        bubbles: true,
        cancelable: true,
        ...opts,
      }),
    );
  });
}

describe("ResizeHandles — primitive contract", () => {
  it("renders all eight handles by default", () => {
    render(<ResizeHandles onHandleGrab={() => {}} />);
    const handles = host.querySelectorAll<HTMLDivElement>("[data-side]");
    const sides = Array.from(handles).map((el) => el.dataset.side);
    expect(sides.sort()).toEqual(["e", "n", "ne", "nw", "s", "se", "sw", "w"]);
  });

  it("respects `sides` to render a subset (e.g. corners-only)", () => {
    const corners: ReadonlyArray<ResizeHandleSide> = ["nw", "ne", "se", "sw"];
    render(<ResizeHandles sides={corners} onHandleGrab={() => {}} />);
    const handles = host.querySelectorAll<HTMLDivElement>("[data-side]");
    expect(handles.length).toBe(4);
    for (const el of Array.from(handles)) {
      expect(isCornerSide(el.dataset.side as ResizeHandleSide)).toBe(true);
      expect(el.dataset.corner).toBe("true");
    }
  });

  it("flags edge handles via data-corner=false", () => {
    render(<ResizeHandles onHandleGrab={() => {}} />);
    const edge = host.querySelector<HTMLDivElement>("[data-side='n']");
    expect(edge?.dataset.corner).toBe("false");
    const corner = host.querySelector<HTMLDivElement>("[data-side='nw']");
    expect(corner?.dataset.corner).toBe("true");
  });

  it("emits onHandleGrab with side / isCorner / coords / shiftKey", () => {
    const grabs: ResizeHandleGrabInfo[] = [];
    render(<ResizeHandles onHandleGrab={(info) => grabs.push(info)} />);

    const seHandle = host.querySelector<HTMLDivElement>("[data-side='se']")!;
    fireMouseDown(seHandle, { clientX: 120, clientY: 200, shiftKey: true });

    expect(grabs).toEqual([
      {
        side: "se",
        isCorner: true,
        shiftKey: true,
        clientX: 120,
        clientY: 200,
      },
    ]);

    const eHandle = host.querySelector<HTMLDivElement>("[data-side='e']")!;
    fireMouseDown(eHandle, { clientX: 10, clientY: 20, shiftKey: false });
    expect(grabs[1]).toEqual({
      side: "e",
      isCorner: false,
      shiftKey: false,
      clientX: 10,
      clientY: 20,
    });
  });

  it("calls preventDefault and stopPropagation on the underlying mousedown", () => {
    // Wrap the primitive in an outer div whose onMouseDown should
    // never see the event because of stopPropagation. We also
    // assert the event is defaultPrevented so consumers don't get
    // a stray text-selection / drag start.
    const outerHandler = vi.fn();
    render(
      <div onMouseDown={outerHandler}>
        <ResizeHandles onHandleGrab={() => {}} />
      </div>,
    );
    const handle = host.querySelector<HTMLDivElement>("[data-side='nw']")!;
    const evt = new MouseEvent("mousedown", { bubbles: true, cancelable: true });
    act(() => {
      handle.dispatchEvent(evt);
    });
    expect(outerHandler).not.toHaveBeenCalled();
    expect(evt.defaultPrevented).toBe(true);
  });

  it("threads dataTestIdPrefix into per-handle data-testid", () => {
    render(<ResizeHandles dataTestIdPrefix="img-foo" onHandleGrab={() => {}} />);
    expect(host.querySelector("[data-testid='img-foo-nw']")).not.toBeNull();
    expect(host.querySelector("[data-testid='img-foo-se']")).not.toBeNull();
  });

  it("uses ns-resize / ew-resize / nwse-resize / nesw-resize cursors", () => {
    render(<ResizeHandles onHandleGrab={() => {}} />);
    const cursorOf = (side: ResizeHandleSide) =>
      host.querySelector<HTMLDivElement>(`[data-side='${side}']`)!.style.cursor;
    expect(cursorOf("n")).toBe("ns-resize");
    expect(cursorOf("s")).toBe("ns-resize");
    expect(cursorOf("e")).toBe("ew-resize");
    expect(cursorOf("w")).toBe("ew-resize");
    expect(cursorOf("nw")).toBe("nwse-resize");
    expect(cursorOf("se")).toBe("nwse-resize");
    expect(cursorOf("ne")).toBe("nesw-resize");
    expect(cursorOf("sw")).toBe("nesw-resize");
  });
});
