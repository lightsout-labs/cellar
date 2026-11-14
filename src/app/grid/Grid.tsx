import {
  useEffect,
  useLayoutEffect,
  useRef,
  useSyncExternalStore,
  type KeyboardEvent,
  type MouseEvent,
} from 'react';
import {
  cellName,
  columnName,
  type Range,
  type Workbook,
} from '../../engine/workbook';
import {
  cellRect,
  COLUMN_HEADER_HEIGHT,
  COLUMN_WIDTH,
  hitTest,
  ROW_HEADER_WIDTH,
  ROW_HEIGHT,
  scrollToShow,
  sheetSize,
  visibleRange,
  type CellPosition,
  type Target,
  type Viewport,
} from './geometry';
import styles from './Grid.module.css';
import { navigationFor } from './navigation';
import { pick } from './selection';
import type { View, ViewState } from './view';

const TEXT_PADDING = 4;
const FONT_SIZE = 13;
const MAC = /Mac|iPhone|iPad/.test(navigator.platform);

/**
 * The cells of the sheet that the view shows, drawn on a canvas that stays
 * in place while the element around it scrolls over the whole sheet. The
 * mouse and the keyboard select cells as in Excel; typing starts an entry,
 * Enter commits it and moves down, and Escape abandons it.
 */
export function Grid({ workbook, view }: { workbook: Workbook; view: View }) {
  const scroller = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  const drag = useRef<Target['kind'] | null>(null);
  const state = useSyncExternalStore(view.subscribe, () => view.state);
  const { sheet, selection, entry, zoom } = state;
  const { active } = selection;
  const activeDisplay = useSyncExternalStore(
    workbook.subscribe.bind(workbook),
    () => workbook.cell(sheet, active.row, active.col).display,
  );

  const viewport = (): Viewport => ({
    scrollX: scroller.current!.scrollLeft,
    scrollY: scroller.current!.scrollTop,
    zoom: view.state.zoom,
  });
  const size = () => ({
    width: scroller.current!.clientWidth,
    height: scroller.current!.clientHeight,
  });
  const topLeft = (): CellPosition => {
    const { width, height } = size();
    const { top, left } = visibleRange(width, height, viewport());
    return { row: top, col: left };
  };
  const show = (cell: CellPosition) => {
    const { width, height } = size();
    const { scrollX, scrollY } = scrollToShow(cell, viewport(), width, height);
    scroller.current!.scrollTo(scrollX, scrollY);
  };

  useEffect(() => {
    const element = scroller.current!;
    const paint = () => {
      const { width, height } = size();
      draw(canvas.current!, workbook, view.state, viewport(), width, height);
    };
    paint();
    const unsubscribeWorkbook = workbook.subscribe(paint);
    const unsubscribeView = view.subscribe(paint);
    // Drawn as soon as the position changes, which is once a frame at most.
    element.addEventListener('scroll', paint);
    const resize = new ResizeObserver(paint);
    resize.observe(element);
    // The colors are those of the theme, which is set on the root element.
    const theme = new MutationObserver(paint);
    theme.observe(document.documentElement, {
      attributeFilter: ['data-theme'],
    });
    return () => {
      unsubscribeWorkbook();
      unsubscribeView();
      element.removeEventListener('scroll', paint);
      resize.disconnect();
      theme.disconnect();
    };
    // `size` and `viewport` read the elements and the view when called.
  }, [workbook, view]);

  // A new zoom keeps the same cell at the top left.
  const shownZoom = useRef(zoom);
  useLayoutEffect(() => {
    const element = scroller.current!;
    const ratio = zoom / shownZoom.current;
    shownZoom.current = zoom;
    element.scrollTo(element.scrollLeft * ratio, element.scrollTop * ratio);
  }, [zoom]);

  // Dragging goes on outside the canvas, and ends wherever the button is
  // released.
  useEffect(() => {
    const onMove = (event: globalThis.MouseEvent) => {
      if (!drag.current || drag.current === 'all') return;
      const bounds = canvas.current!.getBoundingClientRect();
      const { zoom } = view.state;
      let x = event.clientX - bounds.left;
      let y = event.clientY - bounds.top;
      // Over a header, a cell drag takes the cell beside it; a row drag
      // stays in the row header and a column drag in the column header.
      x = drag.current === 'row' ? 0 : Math.max(x, ROW_HEADER_WIDTH * zoom);
      y =
        drag.current === 'column'
          ? 0
          : Math.max(y, COLUMN_HEADER_HEIGHT * zoom);
      const target = hitTest(x, y, viewport());
      const next = pick(view.state.selection, target, 'extend', topLeft());
      if (
        next.focus.row === view.state.selection.focus.row &&
        next.focus.col === view.state.selection.focus.col
      ) {
        return;
      }
      view.select(next);
      show(next.focus);
    };
    const onUp = () => {
      drag.current = null;
    };
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
    return () => {
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    };
  }, [view]);

  const onMouseDown = (event: MouseEvent) => {
    if (event.button !== 0) return;
    const bounds = canvas.current!.getBoundingClientRect();
    const target = hitTest(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
      viewport(),
    );
    const mode = event.shiftKey
      ? 'extend'
      : (MAC ? event.metaKey : event.ctrlKey)
        ? 'add'
        : 'replace';
    view.select(pick(view.state.selection, target, mode, topLeft()));
    drag.current = target.kind;
    // Keep the keyboard on the grid, which the browser would otherwise
    // take away from the editor that is closing.
    event.preventDefault();
    scroller.current!.focus();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (entry !== null) {
      if (event.key === 'Enter') {
        view.navigate({ type: 'step', rows: 1, cols: 0, extend: false });
        show(view.state.selection.active);
      } else if (event.key === 'Escape') {
        view.cancel();
      } else {
        return;
      }
      event.preventDefault();
      scroller.current!.focus();
      return;
    }
    const { width, height } = size();
    const page = {
      rows: Math.max(
        1,
        Math.floor((height / zoom - COLUMN_HEADER_HEIGHT) / ROW_HEIGHT),
      ),
      cols: Math.max(
        1,
        Math.floor((width / zoom - ROW_HEADER_WIDTH) / COLUMN_WIDTH),
      ),
    };
    const navigation = navigationFor(
      {
        key: event.key,
        shift: event.shiftKey,
        ctrl: event.ctrlKey,
        meta: event.metaKey,
        alt: event.altKey,
      },
      { mac: MAC, endMode: state.endMode, page },
    );
    if (navigation) {
      event.preventDefault();
      view.navigate(navigation);
      if (navigation.type === 'endMode' || navigation.type === 'all') return;
      // A page key turns the page, so the cell keeps its place on screen.
      if (event.key === 'PageDown' || event.key === 'PageUp') {
        scroller.current!.scrollBy(
          navigation.type === 'step'
            ? navigation.cols * COLUMN_WIDTH * zoom
            : 0,
          navigation.type === 'step' ? navigation.rows * ROW_HEIGHT * zoom : 0,
        );
      }
      const { selection } = view.state;
      show(navigation.extend ? selection.focus : selection.active);
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

  const sheetArea = sheetSize(zoom);
  const editor = cellRect(active, { scrollX: 0, scrollY: 0, zoom });
  return (
    <div
      ref={scroller}
      className={styles.grid}
      tabIndex={0}
      onKeyDown={onKeyDown}
    >
      <div style={{ width: sheetArea.width, height: sheetArea.height }}>
        <canvas
          ref={canvas}
          className={styles.canvas}
          onMouseDown={onMouseDown}
        />
      </div>
      {entry !== null && (
        <input
          className={styles.editor}
          style={{
            left: editor.x,
            top: editor.y,
            width: editor.width + 1,
            height: editor.height + 1,
            fontSize: FONT_SIZE * zoom,
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
  { sheet, selection }: ViewState,
  viewport: Viewport,
  width: number,
  height: number,
): void {
  const scale = window.devicePixelRatio;
  if (canvas.width !== Math.round(width * scale)) {
    canvas.width = Math.round(width * scale);
    canvas.style.width = `${width}px`;
  }
  if (canvas.height !== Math.round(height * scale)) {
    canvas.height = Math.round(height * scale);
    canvas.style.height = `${height}px`;
  }
  const context = canvas.getContext('2d')!;
  context.setTransform(scale, 0, 0, scale, 0, 0);
  const style = getComputedStyle(canvas);
  const color = (name: string) => style.getPropertyValue(name);
  const { zoom } = viewport;
  const headerWidth = ROW_HEADER_WIDTH * zoom;
  const headerHeight = COLUMN_HEADER_HEIGHT * zoom;
  const range = visibleRange(width, height, viewport);
  const { ranges, active } = selection;
  const rect = (cell: CellPosition) => cellRect(cell, viewport);
  const rangeRect = ({ top, left, bottom, right }: Range) => {
    const start = rect({ row: top, col: left });
    const end = rect({ row: bottom, col: right });
    return {
      x: start.x,
      y: start.y,
      width: end.x + end.width - start.x,
      height: end.y + end.height - start.y,
    };
  };
  const font = `${FONT_SIZE * zoom}px ${style.fontFamily}`;

  context.fillStyle = color('--background');
  context.fillRect(0, 0, width, height);

  // The cells, kept out of the headers.
  context.save();
  context.beginPath();
  context.rect(headerWidth, headerHeight, width, height);
  context.clip();

  context.fillStyle = color('--selection');
  for (const each of ranges) {
    const box = rangeRect(each);
    context.fillRect(box.x, box.y, box.width, box.height);
  }
  if (ranges.length > 1 || !isSingleCell(ranges[0]!)) {
    // The active cell is not shaded, as in Excel.
    const box = rect(active);
    context.fillStyle = color('--background');
    context.fillRect(box.x, box.y, box.width, box.height);
  }

  context.strokeStyle = color('--border');
  context.lineWidth = 1;
  context.beginPath();
  for (let row = range.top; row <= range.bottom + 1; row++) {
    const y = Math.round(rect({ row, col: 0 }).y) + 0.5;
    context.moveTo(0, y);
    context.lineTo(width, y);
  }
  for (let col = range.left; col <= range.right + 1; col++) {
    const x = Math.round(rect({ row: 0, col }).x) + 0.5;
    context.moveTo(x, 0);
    context.lineTo(x, height);
  }
  context.stroke();

  context.font = font;
  context.textBaseline = 'middle';
  context.textAlign = 'left';
  context.fillStyle = color('--text');
  for (let row = range.top; row <= range.bottom; row++) {
    for (let col = range.left; col <= range.right; col++) {
      const { display } = workbook.cell(sheet, row, col);
      if (display === '') continue;
      const box = rect({ row, col });
      context.save();
      context.beginPath();
      context.rect(box.x + 1, box.y + 1, box.width - 1, box.height - 1);
      context.clip();
      context.fillText(
        display,
        box.x + TEXT_PADDING * zoom,
        box.y + box.height / 2 + zoom,
      );
      context.restore();
    }
  }

  context.strokeStyle = color('--accent');
  context.lineWidth = 2;
  if (ranges.length === 1) {
    const box = rangeRect(ranges[0]!);
    context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
  } else {
    const box = rect(active);
    context.lineWidth = 1;
    context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
  }
  context.restore();

  // The headers, those of selected rows and columns marked.
  const covers = (each: Range, row: number | null, col: number | null) =>
    (row === null || (row >= each.top && row <= each.bottom)) &&
    (col === null || (col >= each.left && col <= each.right));
  context.fillStyle = color('--surface');
  context.fillRect(0, 0, width, headerHeight);
  context.fillRect(0, 0, headerWidth, height);
  context.fillStyle = color('--accent-soft');
  for (let row = range.top; row <= range.bottom; row++) {
    if (!ranges.some((each) => covers(each, row, null))) continue;
    const { y, height: h } = rect({ row, col: 0 });
    context.fillRect(0, Math.max(y, headerHeight), headerWidth, h);
  }
  for (let col = range.left; col <= range.right; col++) {
    if (!ranges.some((each) => covers(each, null, col))) continue;
    const { x, width: w } = rect({ row: 0, col });
    context.fillRect(Math.max(x, headerWidth), 0, w, headerHeight);
  }

  context.strokeStyle = color('--border');
  context.lineWidth = 1;
  context.beginPath();
  for (let row = range.top; row <= range.bottom + 1; row++) {
    const y = Math.round(rect({ row, col: 0 }).y) + 0.5;
    if (y < headerHeight) continue;
    context.moveTo(0, y);
    context.lineTo(headerWidth, y);
  }
  for (let col = range.left; col <= range.right + 1; col++) {
    const x = Math.round(rect({ row: 0, col }).x) + 0.5;
    if (x < headerWidth) continue;
    context.moveTo(x, 0);
    context.lineTo(x, headerHeight);
  }
  context.moveTo(0, Math.round(headerHeight) + 0.5);
  context.lineTo(width, Math.round(headerHeight) + 0.5);
  context.moveTo(Math.round(headerWidth) + 0.5, 0);
  context.lineTo(Math.round(headerWidth) + 0.5, height);
  context.stroke();

  context.textAlign = 'center';
  context.fillStyle = color('--muted');
  context.save();
  context.beginPath();
  context.rect(0, headerHeight, headerWidth, height);
  context.clip();
  for (let row = range.top; row <= range.bottom; row++) {
    const { y, height: h } = rect({ row, col: 0 });
    context.fillText(String(row + 1), headerWidth / 2, y + h / 2);
  }
  context.restore();
  context.save();
  context.beginPath();
  context.rect(headerWidth, 0, width, headerHeight);
  context.clip();
  for (let col = range.left; col <= range.right; col++) {
    const { x, width: w } = rect({ row: 0, col });
    context.fillText(columnName(col), x + w / 2, headerHeight / 2);
  }
  context.restore();
}

function isSingleCell({ top, left, bottom, right }: Range): boolean {
  return top === bottom && left === right;
}
