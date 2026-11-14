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
  cellAt,
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
import { formulaHelp, type Highlight } from './formula';
import { FormulaEditor } from './FormulaEditor';
import styles from './Grid.module.css';
import { keyAction } from './keys';
import { pick } from './selection';
import { rowTexts, TEXT_PADDING } from './text';
import type { Entry, View, ViewState } from './view';

const FONT_SIZE = 13;
export const MAC = /Mac|iPhone|iPad/.test(navigator.platform);

/**
 * The cells of the sheet that the view shows, drawn on a canvas that stays
 * in place while the element around it scrolls over the whole sheet. The
 * mouse and the keyboard select and edit cells as in Excel.
 */
export function Grid({ workbook, view }: { workbook: Workbook; view: View }) {
  const scroller = useRef<HTMLDivElement>(null);
  const canvas = useRef<HTMLCanvasElement>(null);
  /** What a drag selects, or the corner of the range it picks for a formula. */
  const drag = useRef<Target['kind'] | CellPosition | null>(null);
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

  // Another sheet opens at its top left, ready for the keyboard unless a
  // formula is being typed in the formula bar.
  const shownSheet = useRef(sheet);
  useLayoutEffect(() => {
    if (shownSheet.current === sheet) return;
    shownSheet.current = sheet;
    scroller.current!.scrollTo(0, 0);
    if (!view.state.entry) scroller.current!.focus();
  }, [sheet, view]);

  // When an entry in the formula bar ends, the keyboard comes back to the
  // grid, which shows the cell it moved to.
  useEffect(() => {
    let previous: Entry | null = view.state.entry;
    return view.subscribe(() => {
      const { entry, selection } = view.state;
      if (previous?.place === 'formulaBar' && entry === null) {
        scroller.current!.focus();
        show(selection.active);
      }
      previous = entry;
    });
    // `show` reads the elements and the view when called.
  }, [view]);

  // Dragging goes on outside the canvas, and ends wherever the button is
  // released.
  useEffect(() => {
    const onMove = (event: globalThis.MouseEvent) => {
      if (!drag.current || drag.current === 'all') return;
      const bounds = canvas.current!.getBoundingClientRect();
      // Past the headers or the edges, the drag reaches the cells beyond
      // them, and the grid scrolls to show them.
      const { row, col } = cellAt(
        event.clientX - bounds.left,
        event.clientY - bounds.top,
        viewport(),
      );
      if (typeof drag.current === 'object') {
        const corner = drag.current;
        view.pick({
          top: Math.min(corner.row, row),
          left: Math.min(corner.col, col),
          bottom: Math.max(corner.row, row),
          right: Math.max(corner.col, col),
        });
        return;
      }
      const target: Target =
        drag.current === 'row'
          ? { kind: 'row', row }
          : drag.current === 'column'
            ? { kind: 'column', col }
            : { kind: 'cell', row, col };
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
    if (target.kind === 'cell' && view.canPick) {
      const { row, col } = target;
      view.pick({ top: row, left: col, bottom: row, right: col });
      drag.current = { row, col };
      // The editor keeps the keyboard.
      event.preventDefault();
      return;
    }
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

  const onDoubleClick = (event: MouseEvent) => {
    const bounds = canvas.current!.getBoundingClientRect();
    const target = hitTest(
      event.clientX - bounds.left,
      event.clientY - bounds.top,
      viewport(),
    );
    if (target.kind === 'cell') view.edit('cell');
  };

  const onKeyDown = (event: KeyboardEvent) => {
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
    const action = keyAction(
      {
        key: event.key,
        shift: event.shiftKey,
        ctrl: event.ctrlKey,
        meta: event.metaKey,
        alt: event.altKey,
      },
      { mac: MAC, mode: entry?.mode ?? null, endMode: state.endMode, page },
    );
    if (!action) return;
    event.preventDefault();
    if (action.type === 'type') {
      // The editor opens holding this first character.
      view.type(action.text);
    } else if (action.type === 'edit') {
      view.edit('cell');
    } else if (action.type === 'clear') {
      view.clear();
    } else if (action.type === 'cycleReference') {
      view.cycleReference();
    } else if (action.type === 'cancel') {
      view.cancel();
      scroller.current!.focus();
    } else {
      const { navigation } = action;
      // The editor closes, and the keyboard stays on the grid.
      scroller.current!.focus();
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
          onDoubleClick={onDoubleClick}
        />
      </div>
      {entry !== null && entry.sheet === sheet && (
        <FormulaEditor
          view={view}
          place="cell"
          label={cellName(active.row, active.col)}
          className={styles.editorBox}
          inputClassName={styles.editor}
          mirrorClassName={styles.editorMirror}
          style={{
            left: editor.x,
            top: editor.y,
            width: editor.width + 1,
            height: editor.height + 1,
            fontSize: FONT_SIZE * zoom,
          }}
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
  { sheet, selection, entry }: ViewState,
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

  // While a formula picks references on another sheet, the selection is
  // that of the formula's sheet.
  const selected = entry === null || entry.sheet === sheet;
  context.fillStyle = color('--selection');
  for (const each of selected ? ranges : []) {
    const box = rangeRect(each);
    context.fillRect(box.x, box.y, box.width, box.height);
  }
  if (selected && (ranges.length > 1 || !isSingleCell(ranges[0]!))) {
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
  context.fillStyle = color('--text');
  const measure = (text: string) => context.measureText(text).width / zoom;
  const padding = TEXT_PADDING * zoom;
  for (let row = range.top; row <= range.bottom; row++) {
    const texts = rowTexts(
      (col) => workbook.cell(sheet, row, col),
      range.left,
      range.right,
      measure,
    );
    for (const { col, text, align, from, to } of texts) {
      const box = rect({ row, col });
      const start = rect({ row, col: from }).x;
      const end = rect({ row, col: to + 1 }).x;
      context.save();
      context.beginPath();
      context.rect(start + 1, box.y + 1, end - start - 1, box.height - 1);
      context.clip();
      // The grid lines inside text that runs over are hidden, as in Excel.
      for (let each = from + 1; each <= to; each++) {
        const x = Math.round(rect({ row, col: each }).x);
        const shaded = (c: number) =>
          selected &&
          ranges.some(
            (r) =>
              row >= r.top && row <= r.bottom && c >= r.left && c <= r.right,
          );
        context.fillStyle = color(
          shaded(each - 1) && shaded(each) ? '--selection' : '--background',
        );
        context.fillRect(x, box.y + 1, 1, box.height - 1);
      }
      context.fillStyle = color('--text');
      context.textAlign = align;
      const x =
        align === 'left'
          ? box.x + padding
          : align === 'right'
            ? box.x + box.width - padding
            : box.x + box.width / 2;
      context.fillText(text, x, box.y + box.height / 2 + zoom);
      context.restore();
    }
  }

  // The ranges of the formula being typed, each in its reference's color.
  context.lineWidth = 2;
  for (const { area, color: index } of shownReferences(
    workbook,
    sheet,
    entry,
  )) {
    const box = rangeRect(area);
    context.strokeStyle = color(`--reference-${index}`);
    context.strokeRect(box.x + 1, box.y + 1, box.width - 1, box.height - 1);
  }

  context.strokeStyle = color('--accent');
  context.lineWidth = 2;
  if (selected && ranges.length === 1) {
    const box = rangeRect(ranges[0]!);
    context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
  } else if (selected) {
    const box = rect(active);
    context.lineWidth = 1;
    context.strokeRect(box.x + 0.5, box.y + 0.5, box.width, box.height);
  }
  context.restore();

  // The headers, those of selected rows and columns marked.
  const covers = (each: Range, row: number | null, col: number | null) =>
    selected &&
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

/** The references of the entry to cells of the shown sheet. */
function shownReferences(
  workbook: Workbook,
  sheet: string,
  entry: Entry | null,
): Highlight[] {
  if (!entry) return [];
  const name = workbook.sheets
    .find((each) => each.id === sheet)!
    .name.toLowerCase();
  return formulaHelp(entry.text, entry.caret).references.filter((reference) =>
    reference.sheet === undefined
      ? entry.sheet === sheet
      : reference.sheet.toLowerCase() === name,
  );
}

function isSingleCell({ top, left, bottom, right }: Range): boolean {
  return top === bottom && left === right;
}
