import * as React from "react";
import type { Slide, SlideSize, ThemeColorScheme } from "../../model/types.js";
import { SlideThumbnail } from "./SlideThumbnail.js";

export interface SlidesSidebarProps {
  readonly slides: ReadonlyArray<Slide>;
  readonly slideSize: SlideSize;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly theme?: ThemeColorScheme;
  readonly activeIndex: number;
  readonly onSelect: (index: number) => void;
  readonly thumbnailWidth?: number;
}

export function SlidesSidebar(props: SlidesSidebarProps): React.ReactElement {
  const { slides, slideSize, mediaUrls, theme, activeIndex, onSelect, thumbnailWidth = 180 } = props;
  return (
    <ul
      className="officeai-pptx-sidebar"
      style={{
        listStyle: "none",
        margin: 0,
        padding: 8,
        display: "flex",
        flexDirection: "column",
        gap: 8,
        overflowY: "auto",
        height: "100%",
      }}
    >
      {slides.map((s, i) => (
        <li key={s.id} style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 12, color: "#71717a", width: 18, textAlign: "right" }}>
            {i + 1}
          </span>
          <SlideThumbnail
            slide={s}
            slideSize={slideSize}
            mediaUrls={mediaUrls}
            theme={theme}
            width={thumbnailWidth}
            active={i === activeIndex}
            onClick={() => onSelect(i)}
            label={`Slide ${i + 1}`}
          />
        </li>
      ))}
    </ul>
  );
}
