import * as RadixTooltip from '@radix-ui/react-tooltip';
import { createContext, useContext, type ReactElement } from 'react';
import { translator, type Key } from './i18n';
import styles from './Shell.module.css';

/** The strings of the language the interface is in. */
export const Translation = createContext<(key: Key) => string>(
  translator('en'),
);

export const useTranslation = () => useContext(Translation);

/** Shows `text` beside the control when it is hovered or focused. */
export function Tooltip({
  text,
  children,
}: {
  text: string;
  children: ReactElement;
}) {
  return (
    <RadixTooltip.Root>
      <RadixTooltip.Trigger asChild>{children}</RadixTooltip.Trigger>
      <RadixTooltip.Portal>
        <RadixTooltip.Content className={styles.tooltip} sideOffset={6}>
          {text}
        </RadixTooltip.Content>
      </RadixTooltip.Portal>
    </RadixTooltip.Root>
  );
}
