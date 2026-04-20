// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { act } from "react";
import { createRoot, type Root } from "react-dom/client";
import type { SheetImage } from "@officeai/xlsx";
import { ImageOverlay } from "./ImageOverlay";
import { I18nProvider } from "../lib/i18n";

// React 19 wants this flag so `act(...)` warnings don't fire on
// every render. Set BEFORE the first `createRoot` call.
(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

/**
 * D8 — pin that the migration to the shared `ResizeHandles`
 * primitive did NOT regress the resize-commit payload that
 * `XlsxEditor` turns into `xlsx:resize-image`. We render the
 * overlay, dispatch a corner-handle mousedown + window
 * mousemove + window mouseup, and assert
 * `onResizeCommit({ widthPx, heightPx })` came back with the
 * expected dimensions.
 *
 * This intentionally exercises the full state-machine including
 * the transient `setTransient` round-trip, NOT just the primitive
 * — that's what makes it an integration test rather than a unit
 * test of the handles themselves.
 */

let host: HTMLDivElement;
let root: Root;

beforeEach(() => {
  host = document.createElement("div");
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

function makeImage(): SheetImage {
  return {
    id: "img-1",
    mediaRef: "media/image1.png",
    name: "test.png",
    anchor: {
      fromRow: 0,
      fromCol: 0,
      fromOffsetXPx: 0,
      fromOffsetYPx: 0,
      widthPx: 100,
      heightPx: 80,
      editAs: "oneCell",
    },
  };
}

function fireMouseDown(el: Element, opts: MouseEventInit): void {
  act(() => {
    el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true, cancelable: true, ...opts }));
  });
}

function fireWindow(type: "mousemove" | "mouseup", opts: MouseEventInit): void {
  act(() => {
    window.dispatchEvent(new MouseEvent(type, { bubbles: true, cancelable: true, ...opts }));
  });
}

describe("ImageOverlay — corner drag commits via onResizeCommit (post-D8)", () => {
  it("dragging the SE handle by (+50, +30) commits widthPx=150, heightPx=110", () => {
    const onResizeCommit = vi.fn();
    const onMoveCommit = vi.fn();
    const onSelect = vi.fn();
    const colXs: number[] = [0, 100, 200, 300, 400];
    const rowYs: number[] = [0, 80, 160, 240, 320];

    act(() => {
      root.render(
        <I18nProvider initialLocale="en">
          <ImageOverlay
            image={makeImage()}
            headerOffset={{ x: 0, y: 0 }}
            colXs={colXs}
            rowYs={rowYs}
            src={undefined}
            selected
            onSelect={onSelect}
            onMoveCommit={onMoveCommit}
            onResizeCommit={onResizeCommit}
            imageId="img-1"
          />
        </I18nProvider>
      );
    });

    const seHandle = host.querySelector("[data-testid='image-handle-img-1-se']");
    expect(seHandle).not.toBeNull();

    fireMouseDown(seHandle!, { clientX: 100, clientY: 80 });
    fireWindow("mousemove", { clientX: 150, clientY: 110 });
    fireWindow("mouseup", { clientX: 150, clientY: 110 });

    expect(onResizeCommit).toHaveBeenCalledTimes(1);
    expect(onResizeCommit).toHaveBeenCalledWith({ widthPx: 150, heightPx: 110 });
    // SE handle never moves the anchor — it pulls only the bottom-right corner.
    expect(onMoveCommit).not.toHaveBeenCalled();
  });

  it("Shift held during a corner drag locks aspect ratio (proportional resize)", () => {
    const onResizeCommit = vi.fn();
    act(() => {
      root.render(
        <I18nProvider initialLocale="en">
          <ImageOverlay
            image={makeImage()}
            headerOffset={{ x: 0, y: 0 }}
            colXs={[0, 100, 200, 300, 400]}
            rowYs={[0, 80, 160, 240, 320]}
            src={undefined}
            selected
            onSelect={() => {}}
            onMoveCommit={() => {}}
            onResizeCommit={onResizeCommit}
            imageId="img-1"
          />
        </I18nProvider>
      );
    });

    const seHandle = host.querySelector("[data-testid='image-handle-img-1-se']")!;
    // Original aspect ratio is 100:80 = 1.25. A +100 / +30 drag
    // would skew the box if we let it; with Shift the height
    // should snap to whichever axis "moved more" (here width
    // dominates: rW=2.0 vs rH=1.375).
    fireMouseDown(seHandle, { clientX: 100, clientY: 80, shiftKey: true });
    fireWindow("mousemove", { clientX: 200, clientY: 110, shiftKey: true });
    fireWindow("mouseup", { clientX: 200, clientY: 110, shiftKey: true });

    expect(onResizeCommit).toHaveBeenCalledTimes(1);
    const call = onResizeCommit.mock.calls[0]![0] as { widthPx: number; heightPx: number };
    expect(call.widthPx).toBe(200);
    // 200 / 1.25 = 160 — locked to the original 1.25 aspect.
    expect(call.heightPx).toBe(160);
  });
});
