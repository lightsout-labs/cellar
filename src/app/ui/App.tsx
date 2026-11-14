import { Provider as TooltipProvider } from '@radix-ui/react-tooltip';
import { useEffect, useLayoutEffect, useState } from 'react';
import type { Workbook } from '../../engine/workbook';
import { Grid, MAC } from '../grid/Grid';
import { View } from '../grid/view';
import type { Platform } from '../platform';
import { restoreWorkbook, startAutosave } from './autosave';
import { Translation } from './controls';
import { startingLanguage, translator, type Language } from './i18n';
import { Ribbon } from './Ribbon';
import { FormulaBar, StatusBar } from './Shell';
import { SheetTabs } from './SheetTabs';
import { shortcut } from './shortcuts';
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
    void platform.workbooks.load(WORKBOOK_ID).then((stored) => {
      if (!current) return;
      const workbook = restoreWorkbook(stored);
      setOpen({ workbook, view: new View(workbook) });
    });
    return () => {
      current = false;
    };
  }, [platform]);

  useEffect(() => {
    if (!workbook) return;
    const autosave = startAutosave(workbook, {
      record: (command) => platform.workbooks.record(WORKBOOK_ID, command),
      save: (snapshot) => platform.workbooks.save(WORKBOOK_ID, snapshot),
    });
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

  // Undo and redo work wherever the focus is, but an entry being typed
  // undoes its own typing.
  useEffect(() => {
    if (!open) return;
    const { view } = open;
    const onKeyDown = (event: KeyboardEvent) => {
      // A text box, such as a sheet name being typed, undoes its own typing.
      if (view.state.entry || isTextBox(event.target)) return;
      const command = shortcut(
        {
          key: event.key,
          shift: event.shiftKey,
          ctrl: event.ctrlKey,
          meta: event.metaKey,
          alt: event.altKey,
        },
        MAC,
      );
      if (!command) return;
      event.preventDefault();
      view[command]();
    };
    window.addEventListener('keydown', onKeyDown);
    return () => window.removeEventListener('keydown', onKeyDown);
  }, [open]);

  if (!open) return null;
  return (
    <Translation value={translator(language)}>
      <TooltipProvider delayDuration={400} disableHoverableContent>
        <div className={styles.app}>
          <Ribbon
            workbook={open.workbook}
            view={open.view}
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

function isTextBox(target: EventTarget | null): boolean {
  return target instanceof HTMLInputElement && target.type === 'text';
}
