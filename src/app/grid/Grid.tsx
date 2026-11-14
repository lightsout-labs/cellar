import {
  useEffect,
  useRef,
  useSyncExternalStore,
  type KeyboardEvent,
  type MouseEvent,
} from 'react';
import { cellName, columnName, type Workbook } from '../../engine/workbook';
import {
  cellAt,
  cellRect,
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
  visibleCells,
  type CellPosition,
} from './geometry';
import styles from './Grid.module.css';
import type { View } from './view';

const TEXT_PADDING = 4;

/**
 * The cells of the sheet that the view shows, drawn on a canvas. A click
 * selects a cell, typing starts an entry in it, Enter commits the entry and
 * moves down, and Escape abandons it.
 */
export function Grid({ workbook, view }: { workbook: Workbook; view: View }) {
  const container = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const { sheet, active, entry } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
  const activeDisplay = useSyncExternalStore(
    workbook.subscribe.bind(workbook),
    () => workbook.cell(sheet, active.row, active.col).display,
  );

  useEffect(() => {
    const paint = () => draw(canvas.current!, workbook, sheet, active);
    paint();
    const unsubscribe = workbook.subscribe(paint);
    const resize = new ResizeObserver(paint);
    resize.observe(canvas.current!);
    // The colors are those of the theme, which is set on the root element.
    const theme = new MutationObserver(paint);
    theme.observe(document.documentElement, {
      attributeFilter: ['data-theme'],
    });
    return () => {
      unsubscribe();
      resize.disconnect();
      theme.disconnect();
    };
  }, [workbook, sheet, active]);

  const onMouseDown = (event: MouseEvent) => {
    const bounds = canvas.current!.getBoundingClientRect();
    const cell = cellAt(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
    );
    if (!cell) return;
    view.select(cell);
    // Keep the keyboard on the grid, which the browser would otherwise
    // take away from the editor that is closing.
    event.preventDefault();
    container.current!.focus();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (entry !== null) {
      if (event.key === 'Enter') {
        view.select({ row: active.row + 1, col: active.col });
      } else if (event.key === 'Escape') {
        view.cancel();
      } else {
        return;
      }
      event.preventDefault();
      container.current!.focus();
    } else if (
      event.key.length === 1 &&
      !event.ctrlKey &&
      !event.metaKey &&
      !event.altKey
    ) {
      // The editor opens holding this first character.
      view.type(event.key);
      event.preventDefault();
    }
  };

  const rect = cellRect(active);
  return (
    <div
      ref={container}
      className={styles.grid}
      tabIndex={0}
      onKeyDown={onKeyDown}
    >
      <canvas
        ref={canvas}
        className={styles.canvas}
        onMouseDown={onMouseDown}
      />
      {entry !== null && (
        <input
          className={styles.editor}
          style={{
            left: rect.x,
            top: rect.y,
            width: rect.width + 1,
            height: rect.height + 1,
          }}
          aria-label={cellName(active.row, active.col)}
          autoFocus
          value={entry}
          onChange={(event) => view.type(event.target.value)}
          onBlur={() => view.commit()}
        />
      )}
      <div role="status" className={styles.announcement}>
        {`${cellName(active.row, active.col)} ${activeDisplay}`.trim()}
      </div>
    </div>
  );
}

function draw(
  canvas: HTMLCanvasElement,
  workbook: Workbook,
  sheet: string,
  active: CellPosition,
): void {
  const width = canvas.clientWidth;
  const height = canvas.clientHeight;
  const scale = window.devicePixelRatio;
  canvas.width = Math.round(width * scale);
  canvas.height = Math.round(height * scale);
  const context = canvas.getContext('2d')!;
  context.scale(scale, scale);
  const style = getComputedStyle(canvas);
  const color = (name: string) => style.getPropertyValue(name);
  const { rows, cols } = visibleCells(width, height);

  context.fillStyle = color('--background');
  context.fillRect(0, 0, width, height);
  context.fillStyle = color('--surface');
  context.fillRect(0, 0, width, COLUMN_HEADER_HEIGHT);
  context.fillRect(0, 0, ROW_HEADER_WIDTH, height);

  context.strokeStyle = color('--border');
  context.lineWidth = 1;
  context.beginPath();
  for (let row = 0; row <= rows; row++) {
    const y = COLUMN_HEADER_HEIGHT + row * ROW_HEIGHT + 0.5;
    context.moveTo(0, y);
    context.lineTo(width, y);
  }
  for (let col = 0; col <= cols; col++) {
    const x = ROW_HEADER_WIDTH + col * COLUMN_WIDTH + 0.5;
    context.moveTo(x, 0);
    context.lineTo(x, height);
  }
  context.stroke();

  context.font = `13px ${style.fontFamily}`;
  context.textBaseline = 'middle';
  context.fillStyle = color('--muted');
  context.textAlign = 'center';
  for (let row = 0; row < rows; row++) {
    const { y } = cellRect({ row, col: 0 });
    context.fillText(String(row + 1), ROW_HEADER_WIDTH / 2, y + ROW_HEIGHT / 2);
  }
  for (let col = 0; col < cols; col++) {
    const { x } = cellRect({ row: 0, col });
    context.fillText(
      columnName(col),
      x + COLUMN_WIDTH / 2,
      COLUMN_HEADER_HEIGHT / 2,
    );
  }

  context.fillStyle = color('--text');
  context.textAlign = 'left';
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      const { display } = workbook.cell(sheet, row, col);
      if (display === '') continue;
      const { x, y } = cellRect({ row, col });
      context.save();
      context.beginPath();
      context.rect(x + 1, y + 1, COLUMN_WIDTH - 1, ROW_HEIGHT - 1);
      context.clip();
      context.fillText(display, x + TEXT_PADDING, y + ROW_HEIGHT / 2 + 1);
      context.restore();
    }
  }

  const rect = cellRect(active);
  context.strokeStyle = color('--accent');
  context.lineWidth = 2;
  context.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width, rect.height);
}
