import { Provider as TooltipProvider } from '@radix-ui/react-tooltip';
import { useEffect, useLayoutEffect, useState } from 'react';
import type { Dropped } from '../../engine/io';
import { Workbook } from '../../engine/workbook';
import { Grid, MAC } from '../grid/Grid';
import { View } from '../grid/view';
import type { Platform } from '../platform';
import { restoreWorkbook, startAutosave } from './autosave';
import { Translation } from './controls';
import { droppedMessage, workbookName } from './files';
import { startingLanguage, translator, type Language } from './i18n';
import { Ribbon } from './Ribbon';
import { FormulaBar, MessageBar, StatusBar } from './Shell';
import { SheetTabs } from './SheetTabs';
import { shortcut } from './shortcuts';
import styles from './Shell.module.css';
import { startingTheme, type Theme } from './theme';

// The workbook shown at the start, until the start screen lists them: the
// last one opened, kept in the preference `workbook`, or else this one.
const FIRST_WORKBOOK_ID = 'default';

/** The workbook shown, under its identifier in storage. */
interface Open {
  readonly id: string;
  readonly workbook: Workbook;
  readonly view: View;
  /** Whether it was just read from a file and is not stored yet. */
  readonly fresh?: boolean;
}

/** What the message bar says about the last file opened. */
type Message =
  | { readonly kind: 'invalid'; readonly name: string }
  | { readonly kind: 'saveFailed' }
  | {
      readonly kind: 'dropped';
      readonly name: string;
      readonly dropped: Dropped[];
    };

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
  const [open, setOpen] = useState<Open>();
  const [message, setMessage] = useState<Message | null>(null);
  const [busy, setBusy] = useState(false);

  // On the root element, so that the tooltips and menus drawn outside the
  // app's own element are themed too, and before the grid first paints.
  useLayoutEffect(() => {
    document.documentElement.lang = language === 'zh' ? 'zh-CN' : 'en';
    document.documentElement.dataset.theme = theme;
  }, [language, theme]);

  useEffect(() => {
    let current = true;
    const id = platform.preferences.get('workbook') ?? FIRST_WORKBOOK_ID;
    void platform.workbooks.load(id).then((stored) => {
      if (!current) return;
      const workbook = restoreWorkbook(stored);
      setOpen({ id, workbook, view: new View(workbook) });
    });
    return () => {
      current = false;
    };
  }, [platform]);

  useEffect(() => {
    if (!open) return;
    const { id, workbook, fresh } = open;
    const autosave = startAutosave(workbook, {
      record: (command) => platform.workbooks.record(id, command),
      save: (snapshot) => platform.workbooks.save(id, snapshot),
    });
    // A workbook read from a file is stored at once, and then shown at the
    // start in place of the one before, which stays stored.
    if (fresh) {
      void platform.workbooks
        .save(id, workbook.toSnapshot())
        .then(() => platform.preferences.set('workbook', id));
    }
    // A hidden tab may be discarded without warning, so it saves at once.
    const onHidden = () => {
      if (document.visibilityState === 'hidden') autosave.flush();
    };
    document.addEventListener('visibilitychange', onHidden);
    return () => {
      document.removeEventListener('visibilitychange', onHidden);
      autosave.stop();
    };
  }, [platform, open]);

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

  const openFile = async () => {
    const file = await platform.openFile();
    if (!file) return;
    setBusy(true);
    try {
      const { readXlsx } = await import('../../engine/io');
      const { snapshot, dropped } = await readXlsx(
        file.bytes,
        workbookName(file.name),
      );
      const workbook = Workbook.fromSnapshot(snapshot);
      setOpen({
        id: crypto.randomUUID(),
        workbook,
        view: new View(workbook),
        fresh: true,
      });
      setMessage(
        dropped.length > 0
          ? { kind: 'dropped', name: file.name, dropped }
          : null,
      );
    } catch {
      // The workbook shown stays as it was.
      setMessage({ kind: 'invalid', name: file.name });
    } finally {
      setBusy(false);
    }
  };

  const saveFile = async (workbook: Workbook) => {
    setBusy(true);
    try {
      const { writeXlsx } = await import('../../engine/io');
      const snapshot = workbook.toSnapshot();
      await platform.saveFile(
        `${snapshot.name}.xlsx`,
        await writeXlsx(snapshot),
      );
    } catch {
      setMessage({ kind: 'saveFailed' });
    } finally {
      setBusy(false);
    }
  };

  if (!open) return null;
  const t = translator(language);
  return (
    <Translation value={t}>
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
            busy={busy}
            onOpen={() => void openFile()}
            onSave={() => void saveFile(open.workbook)}
          />
          {message && (
            <MessageBar
              text={
                message.kind === 'invalid'
                  ? t('file.invalid').replace('{name}', () => message.name)
                  : message.kind === 'saveFailed'
                    ? t('file.saveFailed')
                    : droppedMessage(message.name, message.dropped, language)
              }
              error={message.kind !== 'dropped'}
              onClose={() => setMessage(null)}
            />
          )}
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
