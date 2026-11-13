import { useEffect, useState } from 'react';
import { createWorkbook, Workbook } from '../../engine/workbook';
import { Grid } from '../grid/Grid';
import type { Platform } from '../platform';
import { startAutosave } from './autosave';

// The one workbook there is until the start screen lists several.
const WORKBOOK_ID = 'default';

export function App({ platform }: { platform: Platform }) {
  const [workbook, setWorkbook] = useState<Workbook>();

  useEffect(() => {
    let current = true;
    void platform.workbooks.load(WORKBOOK_ID).then((snapshot) => {
      if (!current) return;
      setWorkbook(
        snapshot ? Workbook.fromSnapshot(snapshot) : createWorkbook(),
      );
    });
    return () => {
      current = false;
    };
  }, [platform]);

  useEffect(() => {
    if (!workbook) return;
    const autosave = startAutosave(workbook, (snapshot) =>
      platform.workbooks.save(WORKBOOK_ID, snapshot),
    );
    // A hidden tab may be discarded without warning, so it saves at once.
    const onHidden = () => {
      if (document.visibilityState === 'hidden') autosave.flush();
    };
    document.addEventListener('visibilitychange', onHidden);
    return () => {
      document.removeEventListener('visibilitychange', onHidden);
      autosave.stop();
    };
  }, [platform, workbook]);

  if (!workbook) return null;
  return <Grid workbook={workbook} sheet={workbook.sheets[0]!.id} />;
}
