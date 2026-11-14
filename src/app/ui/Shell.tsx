import { Minus, Plus } from 'lucide-react';
import { useEffect, useMemo, useReducer, useSyncExternalStore } from 'react';
import { cellName, type Workbook } from '../../engine/workbook';
import { ZOOM_MAX, ZOOM_MIN, type View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import styles from './Shell.module.css';
import { summarize } from './summary';

const ICON_SIZE = 14;
const ZOOM_STEP = 0.1;

/** The name of the active cell and, beside it, what the cell holds. */
export function FormulaBar({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const { sheet, selection, entry } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
  const { active } = selection;
  const contents = useSyncExternalStore(
    workbook.subscribe.bind(workbook),
    () => {
      const cell = workbook.cell(sheet, active.row, active.col);
      return cell.formula ?? cell.display;
    },
  );
  return (
    <div className={styles.formulaLine}>
      <div className={styles.nameBox} role="group" aria-label={t('nameBox')}>
        {cellName(active.row, active.col)}
      </div>
      <div
        className={styles.formulaBar}
        role="group"
        aria-label={t('formulaBar')}
      >
        {entry ?? contents}
      </div>
    </div>
  );
}

/**
 * One tab for each sheet of the workbook, the one shown marked. They only
 * display until there is a second sheet to switch to, so they are a list
 * and not yet controls.
 */
export function SheetTabs({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const { sheet } = useSyncExternalStore(view.subscribe, () => view.state);
  return (
    <div className={styles.sheetTabs} role="list" aria-label={t('sheets')}>
      {workbook.sheets.map(({ id, name }) => (
        <div
          key={id}
          className={styles.sheetTab}
          role="listitem"
          aria-current={id === sheet}
        >
          {name}
        </div>
      ))}
    </div>
  );
}

/**
 * Whether the grid is waiting or an entry is being typed, the average,
 * count and sum of the selected cells, and the zoom.
 */
export function StatusBar({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const { sheet, selection, entry, endMode, zoom } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
  // The figures change with the cells as well as with the selection.
  const [revision, changed] = useReducer((count: number) => count + 1, 0);
  useEffect(() => workbook.subscribe(changed), [workbook]);
  const summary = useMemo(
    () => summarize(workbook, sheet, selection),
    // `revision` stands for the contents of the workbook.
    [workbook, sheet, selection, revision],
  );
  const percent = Math.round(zoom * 100);
  return (
    <footer className={styles.statusBar}>
      <span>{t(entry === null ? 'status.ready' : 'status.enter')}</span>
      {endMode && <span>{t('status.endMode')}</span>}
      <div className={styles.summary} role="group" aria-label={t('summary')}>
        {summary?.average != null && (
          <span>{`${t('summary.average')}: ${summary.average}`}</span>
        )}
        {summary && <span>{`${t('summary.count')}: ${summary.count}`}</span>}
        {summary?.sum != null && (
          <span>{`${t('summary.sum')}: ${summary.sum}`}</span>
        )}
      </div>
      <div className={styles.zoom}>
        <Tooltip text={t('zoom.out.tooltip')}>
          <button
            type="button"
            className={styles.iconButton}
            aria-label={t('zoom.out')}
            disabled={zoom <= ZOOM_MIN}
            onClick={() => view.zoomTo(zoom - ZOOM_STEP)}
          >
            <Minus size={ICON_SIZE} aria-hidden />
          </button>
        </Tooltip>
        <Tooltip text={t('zoom.tooltip')}>
          <input
            type="range"
            className={styles.slider}
            aria-label={t('zoom')}
            min={ZOOM_MIN * 100}
            max={ZOOM_MAX * 100}
            step={10}
            value={percent}
            onChange={(event) => view.zoomTo(Number(event.target.value) / 100)}
          />
        </Tooltip>
        <Tooltip text={t('zoom.in.tooltip')}>
          <button
            type="button"
            className={styles.iconButton}
            aria-label={t('zoom.in')}
            disabled={zoom >= ZOOM_MAX}
            onClick={() => view.zoomTo(zoom + ZOOM_STEP)}
          >
            <Plus size={ICON_SIZE} aria-hidden />
          </button>
        </Tooltip>
        <span className={styles.percent}>{percent}%</span>
      </div>
    </footer>
  );
}
