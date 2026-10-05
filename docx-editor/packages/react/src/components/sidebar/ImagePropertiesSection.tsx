/*
 * Copyright (c) 2026 Casual Office. All rights reserved.
 */

/**
 * Image section of the Format/Properties panel — Google-Docs "Image options"
 * model: text-wrapping as labeled icon tiles (the selected mode highlighted),
 * plus editable width/height. One scannable surface instead of a wall of text
 * options. Reuses the editor's existing wrap + resize commands; this is a
 * presentation layer, not a new command path.
 */
import { useEffect, useState, type CSSProperties, type JSX } from 'react';
import { useTranslation } from '../../i18n';
import type { TranslationKey } from '../../i18n';

const GROUP_HEADER: CSSProperties = {
  padding: '14px 16px 8px',
  fontSize: 11,
  textTransform: 'uppercase',
  letterSpacing: '0.04em',
  color: 'var(--doc-text-muted)',
  fontWeight: 600,
};

const SIZE_ROW: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  gap: 8,
  padding: '4px 16px 6px',
};

const sizeInput: CSSProperties = {
  width: 64,
  padding: '6px 8px',
  fontSize: 13,
  border: '1px solid var(--doc-border, #dadce0)',
  borderRadius: 6,
  background: 'var(--doc-surface, #fff)',
  color: 'var(--doc-text, #202124)',
};

const sizeLabel: CSSProperties = {
  fontSize: 12,
  color: 'var(--doc-text-muted)',
};

export type ImageTransform = 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV';

export interface ImagePropertiesSectionProps {
  /** Current wrap mode of the selected image. */
  wrapType: string;
  /** Current rendered width/height in px (drives the editable inputs). */
  width?: number | null;
  height?: number | null;
  /** Current border attrs (drive the border controls). */
  borderWidth?: number | null;
  borderColor?: string | null;
  /** Current alt text. */
  alt?: string | null;
  /** Apply a wrap mode (host wires this to setImageWrapType). */
  onSetWrap: (value: string) => void;
  /** Apply an explicit width/height (host wires this to setNodeMarkup). */
  onSetSize?: (width: number, height: number) => void;
  /** Rotate/flip the image (host wires this to handleImageTransform). */
  onTransform?: (action: ImageTransform) => void;
  /** Set border width/color/style (host wires this to setNodeMarkup). */
  onSetBorder?: (
    borderWidth: number | null,
    borderColor: string | null,
    borderStyle: string | null
  ) => void;
  /** Set alt text (host wires this to setNodeMarkup). */
  onSetAlt?: (alt: string) => void;
  /** Current distance-from-text margins (px) for wrapped images. */
  distTop?: number | null;
  distBottom?: number | null;
  distLeft?: number | null;
  distRight?: number | null;
  /** Set distance-from-text margins (host wires this to setNodeMarkup). */
  onSetDist?: (side: 'distTop' | 'distBottom' | 'distLeft' | 'distRight', value: number) => void;
}

// Distance-from-text only applies to text-wrapping floats.
const WRAPS_TEXT = new Set(['squareLeft', 'squareRight', 'topAndBottom']);

const BORDER_PRESETS: { id: string; labelKey: TranslationKey; width: number | null }[] = [
  { id: 'none', labelKey: 'dialogs.bordersShading.presetNone', width: null },
  { id: 'thin', labelKey: 'sidebar.imageProperties.borderThin', width: 1 },
  { id: 'medium', labelKey: 'sidebar.imageProperties.borderMedium', width: 2 },
  { id: 'thick', labelKey: 'dialogs.bordersShading.styleThick', width: 4 },
];

const ARRANGE: { action: ImageTransform; labelKey: TranslationKey; icon: JSX.Element }[] = [
  {
    action: 'rotateCCW',
    labelKey: 'sidebar.imageProperties.rotateLeft',
    icon: (
      <svg width="18" height="18" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
        <path d="M7.11 8.53 5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.87-1.59-1.01-2.47zm1.01 5.32c1.16.9 2.51 1.44 3.9 1.61V17.9c-.87-.15-1.71-.49-2.46-1.03L7.1 18.32zM13 4.07V1L8.45 5.55 13 10V6.09c2.84.48 5 2.94 5 5.91s-2.16 5.43-5 5.91v2.02c3.95-.49 7-3.85 7-7.93s-3.05-7.44-7-7.93z" />
      </svg>
    ),
  },
  {
    action: 'rotateCW',
    labelKey: 'sidebar.imageProperties.rotateRight',
    icon: (
      <svg
        width="18"
        height="18"
        viewBox="0 0 24 24"
        fill="currentColor"
        aria-hidden="true"
        style={{ transform: 'scaleX(-1)' }}
      >
        <path d="M7.11 8.53 5.7 7.11C4.8 8.27 4.24 9.61 4.07 11h2.02c.14-.87.49-1.72 1.02-2.47zM6.09 13H4.07c.17 1.39.72 2.73 1.62 3.89l1.41-1.42c-.52-.75-.87-1.59-1.01-2.47zm1.01 5.32c1.16.9 2.51 1.44 3.9 1.61V17.9c-.87-.15-1.71-.49-2.46-1.03L7.1 18.32zM13 4.07V1L8.45 5.55 13 10V6.09c2.84.48 5 2.94 5 5.91s-2.16 5.43-5 5.91v2.02c3.95-.49 7-3.85 7-7.93s-3.05-7.44-7-7.93z" />
      </svg>
    ),
  },
  {
    action: 'flipH',
    labelKey: 'imageTransform.flipHorizontal',
    icon: (
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M12 3v18" stroke="currentColor" strokeWidth="1.6" strokeDasharray="2 2" />
        <path d="M10 6 4 12l6 6V6zM14 6v12l6-6-6-6z" fill="currentColor" opacity="0.85" />
      </svg>
    ),
  },
  {
    action: 'flipV',
    labelKey: 'imageTransform.flipVertical',
    icon: (
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path d="M3 12h18" stroke="currentColor" strokeWidth="1.6" strokeDasharray="2 2" />
        <path d="M6 10 12 4l6 6H6zM6 14h12l-6 6-6-6z" fill="currentColor" opacity="0.85" />
      </svg>
    ),
  },
];

const arrangeBtn: CSSProperties = {
  display: 'inline-flex',
  alignItems: 'center',
  justifyContent: 'center',
  width: 40,
  height: 36,
  border: '1px solid var(--doc-border, #dadce0)',
  borderRadius: 8,
  background: 'transparent',
  color: 'var(--doc-text, #202124)',
  cursor: 'pointer',
};

export function ImagePropertiesSection({
  wrapType,
  width,
  height,
  borderWidth,
  borderColor,
  alt,
  onSetSize,
  onTransform,
  onSetBorder,
  onSetAlt,
  distTop,
  distBottom,
  distLeft,
  distRight,
  onSetDist,
}: ImagePropertiesSectionProps) {
  const { t } = useTranslation();
  // Local, editable copies so typing doesn't fight the live node attrs. They
  // re-sync whenever the selected image (its size) changes underneath.
  const [w, setW] = useState<string>(width != null ? String(Math.round(width)) : '');
  const [h, setH] = useState<string>(height != null ? String(Math.round(height)) : '');
  const [lockAspect, setLockAspect] = useState(true);
  const [color, setColor] = useState<string>(borderColor || '#000000');
  const [altText, setAltText] = useState<string>(alt || '');

  useEffect(() => {
    setW(width != null ? String(Math.round(width)) : '');
    setH(height != null ? String(Math.round(height)) : '');
  }, [width, height]);
  useEffect(() => {
    setColor(borderColor || '#000000');
  }, [borderColor]);
  useEffect(() => {
    setAltText(alt || '');
  }, [alt]);

  const aspect = width && height ? width / height : null;

  const commitWidth = () => {
    const nw = Number(w);
    if (!Number.isFinite(nw) || nw <= 0 || !onSetSize) return;
    const nh = lockAspect && aspect ? Math.round(nw / aspect) : Number(h) || nw;
    setH(String(nh));
    onSetSize(nw, nh);
  };
  const commitHeight = () => {
    const nh = Number(h);
    if (!Number.isFinite(nh) || nh <= 0 || !onSetSize) return;
    const nw = lockAspect && aspect ? Math.round(nh * aspect) : Number(w) || nh;
    setW(String(nw));
    onSetSize(nw, nh);
  };
  const onKey = (e: React.KeyboardEvent, commit: () => void) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      commit();
    }
  };

  return (
    <div data-testid="properties-image-section">
      {onSetSize && (width != null || height != null) && (
        <>
          <div style={GROUP_HEADER}>{t('dialogs.pageSetup.sizeLabel')}</div>
          <div style={SIZE_ROW} data-testid="properties-image-size">
            <label style={sizeLabel}>
              {t('sidebar.imageProperties.widthAbbr')}
              <input
                style={{ ...sizeInput, marginLeft: 6 }}
                type="number"
                min={8}
                max={2000}
                value={w}
                data-testid="properties-image-width"
                onChange={(e) => setW(e.target.value)}
                onBlur={commitWidth}
                onKeyDown={(e) => onKey(e, commitWidth)}
              />
            </label>
            <label style={sizeLabel}>
              {t('sidebar.imageProperties.heightAbbr')}
              <input
                style={{ ...sizeInput, marginLeft: 6 }}
                type="number"
                min={8}
                max={2000}
                value={h}
                data-testid="properties-image-height"
                onChange={(e) => setH(e.target.value)}
                onBlur={commitHeight}
                onKeyDown={(e) => onKey(e, commitHeight)}
              />
            </label>
          </div>
          <label
            style={{
              ...sizeLabel,
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              padding: '0 16px 14px',
              cursor: 'pointer',
            }}
          >
            <input
              type="checkbox"
              checked={lockAspect}
              data-testid="properties-image-lock-aspect"
              onChange={(e) => setLockAspect(e.target.checked)}
            />
            {t('sidebar.imageProperties.lockAspectRatio')}
          </label>
        </>
      )}

      {onSetDist && WRAPS_TEXT.has(wrapType) && (
        <>
          <div style={GROUP_HEADER}>{t('sidebar.imageProperties.distanceFromText')}</div>
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: '1fr 1fr',
              gap: '6px 12px',
              padding: '0 16px 14px',
            }}
            data-testid="properties-image-dist"
          >
            {(
              [
                ['distTop', t('dialogs.bordersShading.side.top'), distTop],
                ['distBottom', t('dialogs.bordersShading.side.bottom'), distBottom],
                ['distLeft', t('dialogs.bordersShading.side.left'), distLeft],
                ['distRight', t('dialogs.bordersShading.side.right'), distRight],
              ] as const
            ).map(([side, label, val]) => (
              <label
                key={side}
                style={{ ...sizeLabel, display: 'flex', alignItems: 'center', gap: 6 }}
              >
                {label}
                <input
                  style={{ ...sizeInput, width: 56 }}
                  type="number"
                  min={0}
                  max={200}
                  defaultValue={val != null ? Math.round(val) : 0}
                  data-testid={`properties-image-${side}`}
                  onBlur={(e) => onSetDist(side, Number(e.target.value) || 0)}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      e.preventDefault();
                      onSetDist(side, Number((e.target as HTMLInputElement).value) || 0);
                    }
                  }}
                />
              </label>
            ))}
          </div>
        </>
      )}

      {onTransform && (
        <>
          <div style={GROUP_HEADER}>{t('sidebar.imageProperties.arrange')}</div>
          <div
            style={{ display: 'flex', gap: 6, padding: '0 16px 4px' }}
            role="group"
            aria-label={t('sidebar.imageProperties.arrange')}
          >
            {ARRANGE.map((a) => {
              const label = t(a.labelKey);
              return (
                <button
                  key={a.action}
                  type="button"
                  title={label}
                  aria-label={label}
                  style={arrangeBtn}
                  data-testid={`properties-image-${a.action}`}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    onTransform(a.action);
                  }}
                >
                  {a.icon}
                </button>
              );
            })}
          </div>
        </>
      )}

      {onSetBorder && (
        <>
          <div style={GROUP_HEADER}>{t('dialogs.imageProperties.border')}</div>
          <div
            style={{ display: 'flex', flexWrap: 'wrap', gap: 6, padding: '0 16px 6px' }}
            role="group"
            aria-label={t('table.borderWidth')}
          >
            {BORDER_PRESETS.map((b) => {
              const active = (borderWidth ?? null) === b.width;
              return (
                <button
                  key={b.id}
                  type="button"
                  style={{
                    ...arrangeBtn,
                    width: 'auto',
                    padding: '0 12px',
                    fontSize: 12.5,
                    color: active ? 'var(--doc-primary, #1a73e8)' : 'var(--doc-text, #202124)',
                    background: active ? 'var(--doc-primary-light, #e8f0fe)' : 'transparent',
                    borderColor: active
                      ? 'var(--doc-primary, #1a73e8)'
                      : 'var(--doc-border, #dadce0)',
                  }}
                  data-testid={`properties-image-border-${b.id}`}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    onSetBorder(
                      b.width,
                      b.width == null ? null : color,
                      b.width == null ? null : 'solid'
                    );
                  }}
                >
                  {t(b.labelKey)}
                </button>
              );
            })}
          </div>
          <label style={{ ...SIZE_ROW, ...sizeLabel, paddingBottom: 14 }}>
            {t('dialogs.imageProperties.color')}
            <input
              type="color"
              value={color}
              data-testid="properties-image-border-color"
              style={{
                marginLeft: 8,
                width: 36,
                height: 26,
                padding: 0,
                border: '1px solid var(--doc-border, #dadce0)',
                borderRadius: 6,
                background: 'none',
                cursor: 'pointer',
              }}
              onChange={(e) => {
                setColor(e.target.value);
                // Apply immediately if a border is already on; else just stage.
                if (borderWidth) onSetBorder(borderWidth, e.target.value, 'solid');
              }}
            />
          </label>
        </>
      )}

      {onSetAlt && (
        <>
          <div style={GROUP_HEADER}>{t('sidebar.imageProperties.altText')}</div>
          <div style={{ padding: '0 16px 16px' }}>
            <textarea
              value={altText}
              data-testid="properties-image-alt"
              placeholder={t('sidebar.imageProperties.altTextPlaceholder')}
              rows={2}
              style={{
                width: '100%',
                resize: 'vertical',
                padding: '6px 8px',
                fontSize: 13,
                border: '1px solid var(--doc-border, #dadce0)',
                borderRadius: 6,
                background: 'var(--doc-surface, #fff)',
                color: 'var(--doc-text, #202124)',
                boxSizing: 'border-box',
              }}
              onChange={(e) => setAltText(e.target.value)}
              onBlur={() => onSetAlt(altText)}
            />
          </div>
        </>
      )}
    </div>
  );
}
