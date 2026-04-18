"use client";

import {
  FileUp,
  Download,
  Plus,
  Trash2,
  Copy,
  Bold,
  Italic,
  Underline,
  Type,
  ZoomIn,
  ZoomOut,
} from "lucide-react";

export interface PptxToolbarProps {
  readonly disabled: boolean;
  readonly slideCount: number;
  readonly onOpenFile: () => void;
  readonly onExport: () => void;
  readonly onAddSlide: () => void;
  readonly onDeleteSlide: () => void;
  readonly onDuplicateSlide: () => void;
  readonly onAddTextBox: () => void;
  readonly onToggleBold: () => void;
  readonly onToggleItalic: () => void;
  readonly onToggleUnderline: () => void;
  readonly zoom: number;
  readonly minZoom: number;
  readonly maxZoom: number;
  readonly onZoomChange: (zoom: number) => void;
  readonly onZoomReset: () => void;
}

export function PptxToolbar(props: PptxToolbarProps) {
  const { disabled, zoom, minZoom, maxZoom, onZoomChange, onZoomReset } = props;
  const zoomPct = Math.round(zoom * 100);
  return (
    <div className="flex flex-wrap items-center gap-1 border-b border-divider pb-2">
      <ToolbarButton onClick={props.onOpenFile} icon={<FileUp size={14} />} label="Open" />
      <ToolbarButton onClick={props.onExport} icon={<Download size={14} />} label="Export" disabled={disabled} />
      <span className="mx-2 h-5 w-px bg-divider" />
      <ToolbarButton onClick={props.onAddSlide} icon={<Plus size={14} />} label="Add slide" disabled={disabled} />
      <ToolbarButton
        onClick={props.onDuplicateSlide}
        icon={<Copy size={14} />}
        label="Duplicate"
        disabled={disabled || props.slideCount < 1}
      />
      <ToolbarButton
        onClick={props.onDeleteSlide}
        icon={<Trash2 size={14} />}
        label="Delete"
        disabled={disabled || props.slideCount <= 1}
      />
      <span className="mx-2 h-5 w-px bg-divider" />
      <ToolbarButton onClick={props.onAddTextBox} icon={<Type size={14} />} label="Text box" disabled={disabled} />
      <span className="mx-2 h-5 w-px bg-divider" />
      <ToolbarButton onClick={props.onToggleBold} icon={<Bold size={14} />} label="Bold" disabled={disabled} />
      <ToolbarButton onClick={props.onToggleItalic} icon={<Italic size={14} />} label="Italic" disabled={disabled} />
      <ToolbarButton onClick={props.onToggleUnderline} icon={<Underline size={14} />} label="Underline" disabled={disabled} />
      <span className="mx-2 h-5 w-px bg-divider" />
      <div className="ml-auto flex items-center gap-2 text-xs text-secondary">
        <ToolbarButton
          onClick={() => onZoomChange(Math.max(minZoom, zoom - 0.1))}
          icon={<ZoomOut size={14} />}
          label="Zoom out"
          disabled={disabled || zoom <= minZoom + 1e-6}
        />
        <input
          type="range"
          aria-label="Zoom"
          data-testid="pptx-zoom-slider"
          min={minZoom}
          max={maxZoom}
          step={0.05}
          value={zoom}
          disabled={disabled}
          onChange={(e) => onZoomChange(Number(e.target.value))}
          className="h-1 w-28 accent-[var(--accent)]"
        />
        <ToolbarButton
          onClick={() => onZoomChange(Math.min(maxZoom, zoom + 0.1))}
          icon={<ZoomIn size={14} />}
          label="Zoom in"
          disabled={disabled || zoom >= maxZoom - 1e-6}
        />
        <button
          type="button"
          onClick={onZoomReset}
          disabled={disabled}
          title="Reset zoom"
          data-testid="pptx-zoom-reset"
          className="min-w-[44px] rounded px-1.5 py-0.5 text-xs tabular-nums text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
        >
          {zoomPct}%
        </button>
      </div>
    </div>
  );
}

interface ToolbarButtonProps {
  readonly onClick: () => void;
  readonly icon: React.ReactNode;
  readonly label: string;
  readonly disabled?: boolean;
}

function ToolbarButton({ onClick, icon, label, disabled }: ToolbarButtonProps) {
  return (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled}
      title={label}
      className="inline-flex items-center gap-1.5 rounded px-2 py-1 text-xs text-foreground hover:bg-hover disabled:cursor-not-allowed disabled:opacity-40"
    >
      {icon}
      <span className="hidden sm:inline">{label}</span>
    </button>
  );
}
