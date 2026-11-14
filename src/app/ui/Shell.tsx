import { useSyncExternalStore } from 'react';
import { cellName, type Workbook } from '../../engine/workbook';
import type { View } from '../grid/view';
import { useTranslation } from './controls';
import styles from './Shell.module.css';

/** The name of the active cell and, beside it, what the cell holds. */
export function FormulaBar({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const { sheet, active, entry } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
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

/** Whether the grid is waiting or an entry is being typed. */
export function StatusBar({ view }: { view: View }) {
  const t = useTranslation();
  const { entry } = useSyncExternalStore(view.subscribe, () => view.state);
  return (
    <footer className={styles.statusBar}>
      {t(entry === null ? 'status.ready' : 'status.enter')}
    </footer>
  );
}
