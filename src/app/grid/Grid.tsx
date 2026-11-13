import {
  useEffect,
  useRef,
  useState,
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

const COLORS = {
  background: '#ffffff',
  header: '#f3f3f3',
  line: '#e1e1e1',
  headerText: '#5f5f5f',
  text: '#1f1f1f',
  accent: '#217346',
};
const TEXT_PADDING = 4;

/**
 * The cells of one sheet, drawn on a canvas. A click selects a cell, typing
 * starts an entry in it, Enter commits the entry and moves down, and Escape
 * abandons it.
 */
export function Grid({
  workbook,
  sheet,
}: {
  workbook: Workbook;
  sheet: string;
}) {
  const container = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const [active, setActive] = useState<CellPosition>({ row: 0, col: 0 });
  /** The text of the entry being typed, or `null` when there is none. */
  const [entry, setEntry] = useState<string | null>(null);
  const activeDisplay = useSyncExternalStore(
    workbook.subscribe.bind(workbook),
    () => workbook.cell(sheet, active.row, active.col).display,
  );

  useEffect(() => {
    const paint = () => draw(canvas.current!, workbook, sheet, active);
    paint();
    const unsubscribe = workbook.subscribe(paint);
    const observer = new ResizeObserver(paint);
    observer.observe(canvas.current!);
    return () => {
      unsubscribe();
      observer.disconnect();
    };
  }, [workbook, sheet, active]);

  // Closing the editor takes the keyboard from it, and that blur must not
  // commit an entry that was just committed or abandoned.
  const closed = useRef(true);
  const close = (save: boolean) => {
    if (closed.current) return;
    closed.current = true;
    if (save) {
      workbook.execute({ type: 'setCell', sheet, ...active, input: entry! });
    }
    setEntry(null);
  };

  const onMouseDown = (event: MouseEvent) => {
    const bounds = canvas.current!.getBoundingClientRect();
    const cell = cellAt(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
    );
    if (!cell) return;
    close(true);
    setActive(cell);
    // Keep the keyboard on the grid, which the browser would otherwise
    // take away from the editor that is closing.
    event.preventDefault();
    container.current!.focus();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (entry !== null) {
      if (event.key === 'Enter') {
        close(true);
        setActive({ row: active.row + 1, col: active.col });
      } else if (event.key === 'Escape') {
        close(false);
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
      closed.current = false;
      setEntry(event.key);
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
          onChange={(event) => setEntry(event.target.value)}
          onBlur={() => close(true)}
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
  const { rows, cols } = visibleCells(width, height);

  context.fillStyle = COLORS.background;
  context.fillRect(0, 0, width, height);
  context.fillStyle = COLORS.header;
  context.fillRect(0, 0, width, COLUMN_HEADER_HEIGHT);
  context.fillRect(0, 0, ROW_HEADER_WIDTH, height);

  context.strokeStyle = COLORS.line;
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

  context.font = `13px ${getComputedStyle(canvas).fontFamily}`;
  context.textBaseline = 'middle';
  context.fillStyle = COLORS.headerText;
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

  context.fillStyle = COLORS.text;
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
  context.strokeStyle = COLORS.accent;
  context.lineWidth = 2;
  context.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.width, rect.height);
}
