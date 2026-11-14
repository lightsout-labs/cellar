import * as Tabs from '@radix-ui/react-tabs';
import { Moon, Redo2, Sun, Table2, Undo2, type LucideIcon } from 'lucide-react';
import {
  useCallback,
  useId,
  useSyncExternalStore,
  type ReactNode,
} from 'react';
import type { Workbook } from '../../engine/workbook';
import { MAC } from '../grid/Grid';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import type { Language } from './i18n';
import styles from './Shell.module.css';
import type { Theme } from './theme';

const ICON_SIZE = 16;

/**
 * The toolbar at the top of the window: tabs, each with the controls of
 * its features in labeled groups. A tab or control is added by the slice
 * that builds its feature.
 */
export function Ribbon({
  workbook,
  view,
  language,
  onLanguage,
  theme,
  onTheme,
}: {
  workbook: Workbook;
  view: View;
  language: Language;
  onLanguage: (language: Language) => void;
  theme: Theme;
  onTheme: (theme: Theme) => void;
}) {
  const t = useTranslation();
  return (
    <Tabs.Root className={styles.ribbon} defaultValue="home">
      <div className={styles.ribbonTabs}>
        <span className={styles.brand}>
          <Table2 size={ICON_SIZE} aria-hidden />
          Gridworks
        </span>
        <Tabs.List aria-label={t('ribbon')}>
          <Tooltip text={t('tab.home.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="home">
              {t('tab.home')}
            </Tabs.Trigger>
          </Tooltip>
          <Tooltip text={t('tab.view.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="view">
              {t('tab.view')}
            </Tabs.Trigger>
          </Tooltip>
        </Tabs.List>
      </div>
      <Tabs.Content className={styles.ribbonPanel} value="home">
        <Group label={t('undo')}>
          <History workbook={workbook} view={view} />
        </Group>
      </Tabs.Content>
      <Tabs.Content className={styles.ribbonPanel} value="view">
        <Group label={t('theme')}>
          <Choice
            icon={Sun}
            label={t('theme.light')}
            tooltip={t('theme.light.tooltip')}
            chosen={theme === 'light'}
            onChoose={() => onTheme('light')}
          />
          <Choice
            icon={Moon}
            label={t('theme.dark')}
            tooltip={t('theme.dark.tooltip')}
            chosen={theme === 'dark'}
            onChoose={() => onTheme('dark')}
          />
        </Group>
        <Group label={t('language')}>
          {/* Each language is named in itself, whatever the interface is in. */}
          <Choice
            label="English"
            tooltip={t('language.en.tooltip')}
            chosen={language === 'en'}
            onChoose={() => onLanguage('en')}
          />
          <Choice
            label="中文"
            tooltip={t('language.zh.tooltip')}
            chosen={language === 'zh'}
            onChoose={() => onLanguage('zh')}
          />
        </Group>
      </Tabs.Content>
    </Tabs.Root>
  );
}

function Group({ label, children }: { label: string; children: ReactNode }) {
  const id = useId();
  return (
    <div className={styles.group} role="group" aria-labelledby={id}>
      <span id={id} className={styles.groupLabel}>
        {label}
      </span>
      {children}
    </div>
  );
}

/** The Undo and Redo buttons, enabled when there is something to do. */
function History({ workbook, view }: { workbook: Workbook; view: View }) {
  const t = useTranslation();
  const subscribe = useCallback(
    (listener: () => void) => workbook.subscribe(listener),
    [workbook],
  );
  const canUndo = useSyncExternalStore(subscribe, () => workbook.canUndo);
  const canRedo = useSyncExternalStore(subscribe, () => workbook.canRedo);
  const key = MAC ? '⌘' : 'Ctrl+';
  return (
    <>
      <Command
        icon={Undo2}
        label={t('undo')}
        tooltip={`${t('undo.tooltip')} (${key}Z)`}
        disabled={!canUndo}
        onRun={() => view.undo()}
      />
      <Command
        icon={Redo2}
        label={t('redo')}
        tooltip={`${t('redo.tooltip')} (${key}Y)`}
        disabled={!canRedo}
        onRun={() => view.redo()}
      />
    </>
  );
}

/** A button that runs a command, shown by its icon alone. */
function Command({
  icon: Icon,
  label,
  tooltip,
  disabled,
  onRun,
}: {
  icon: LucideIcon;
  label: string;
  tooltip: string;
  disabled: boolean;
  onRun: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        aria-label={label}
        disabled={disabled}
        onClick={onRun}
      >
        <Icon size={ICON_SIZE} aria-hidden />
      </button>
    </Tooltip>
  );
}

/** A button for one of the alternatives of a setting. */
function Choice({
  icon: Icon,
  label,
  tooltip,
  chosen,
  onChoose,
}: {
  icon?: LucideIcon;
  label: string;
  tooltip: string;
  chosen: boolean;
  onChoose: () => void;
}) {
  return (
    <Tooltip text={tooltip}>
      <button
        type="button"
        className={styles.button}
        aria-pressed={chosen}
        onClick={onChoose}
      >
        {Icon && <Icon size={ICON_SIZE} aria-hidden />}
        {label}
      </button>
    </Tooltip>
  );
}
