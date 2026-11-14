import { Provider as TooltipProvider } from '@radix-ui/react-tooltip';
import { useEffect, useLayoutEffect, useState } from 'react';
import { createWorkbook, Workbook } from '../../engine/workbook';
import { Grid } from '../grid/Grid';
import { View } from '../grid/view';
import type { Platform } from '../platform';
import { startAutosave } from './autosave';
import { Translation } from './controls';
import { startingLanguage, translator, type Language } from './i18n';
import { Ribbon } from './Ribbon';
import { FormulaBar, SheetTabs, StatusBar } from './Shell';
import styles from './Shell.module.css';
import { startingTheme, type Theme } from './theme';

// The one workbook there is until the start screen lists several.
const WORKBOOK_ID = 'default';

export function App({ platform }: { platform: Platform }) {
  const { preferences } = platform;
  const [language, setLanguage] = useState(() =>
    startingLanguage(preferences.get('language'), navigator.language),
  );
  const [theme, setTheme] = useState(() =>
    startingTheme(
      preferences.get('theme'),
      window.matchMedia('(prefers-color-scheme: dark)').matches,
    ),
  );
  const [open, setOpen] = useState<{ workbook: Workbook; view: View }>();
  const workbook = open?.workbook;

  // On the root element, so that the tooltips and menus drawn outside the
  // app's own element are themed too, and before the grid first paints.
  useLayoutEffect(() => {
    document.documentElement.lang = language === 'zh' ? 'zh-CN' : 'en';
    document.documentElement.dataset.theme = theme;
  }, [language, theme]);

  useEffect(() => {
    let current = true;
    void platform.workbooks.load(WORKBOOK_ID).then((snapshot) => {
      if (!current) return;
      const workbook = snapshot
        ? Workbook.fromSnapshot(snapshot)
        : createWorkbook();
      setOpen({ workbook, view: new View(workbook) });
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

  if (!open) return null;
  return (
    <Translation value={translator(language)}>
      <TooltipProvider delayDuration={400} disableHoverableContent>
        <div className={styles.app}>
          <Ribbon
            language={language}
            onLanguage={(language: Language) => {
              setLanguage(language);
              preferences.set('language', language);
            }}
            theme={theme}
            onTheme={(theme: Theme) => {
              setTheme(theme);
              preferences.set('theme', theme);
            }}
          />
          <FormulaBar {...open} />
          <main className={styles.grid}>
            <Grid {...open} />
          </main>
          <SheetTabs {...open} />
          <StatusBar {...open} />
        </div>
      </TooltipProvider>
    </Translation>
  );
}
