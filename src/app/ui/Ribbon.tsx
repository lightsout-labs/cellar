import * as Tabs from '@radix-ui/react-tabs';
import { Moon, Sun, Table2, type LucideIcon } from 'lucide-react';
import { useId, type ReactNode } from 'react';
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
  language,
  onLanguage,
  theme,
  onTheme,
}: {
  language: Language;
  onLanguage: (language: Language) => void;
  theme: Theme;
  onTheme: (theme: Theme) => void;
}) {
  const t = useTranslation();
  return (
    <Tabs.Root className={styles.ribbon} defaultValue="view">
      <div className={styles.ribbonTabs}>
        <span className={styles.brand}>
          <Table2 size={ICON_SIZE} aria-hidden />
          Cellar
        </span>
        <Tabs.List aria-label={t('ribbon')}>
          <Tooltip text={t('tab.view.tooltip')}>
            <Tabs.Trigger className={styles.ribbonTab} value="view">
              {t('tab.view')}
            </Tabs.Trigger>
          </Tooltip>
        </Tabs.List>
      </div>
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
