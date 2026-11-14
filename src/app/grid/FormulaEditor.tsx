import {
  useId,
  useLayoutEffect,
  useRef,
  useState,
  useSyncExternalStore,
  type ComponentPropsWithRef,
  type FocusEvent,
  type KeyboardEvent,
  type ReactNode,
} from 'react';
import { useTranslation } from '../ui/controls';
import styles from './FormulaEditor.module.css';
import { functions } from '../../engine/formula';
import { formulaHelp } from './formula';
import type { Entry, View } from './view';

/** The attribute that marks the editors of an entry, valued with their place. */
export const EDITOR = 'data-entry-editor';

/**
 * Whether focus that leaves an editor goes to the other one, so that the
 * entry goes on rather than being committed.
 */
function staysInEntry(event: FocusEvent): boolean {
  return (
    event.relatedTarget instanceof Element &&
    event.relatedTarget.hasAttribute(EDITOR)
  );
}

/**
 * The box where an entry is typed, in the cell or the formula bar, or
 * where the formula bar shows `contents` when there is no entry. A formula
 * shows each reference in its color, the functions whose names start
 * with the one typed, and the arguments of the function the caret is in.
 * Other props go to the element around the box, as a tooltip gives them.
 */
export function FormulaEditor({
  view,
  place,
  contents = '',
  label,
  className,
  inputClassName,
  mirrorClassName,
  onKeyDown,
  ...around
}: {
  view: View;
  place: Entry['place'];
  contents?: string;
  label: string;
  className?: string;
  inputClassName?: string;
  /** Lays the colored text exactly over the text of the box. */
  mirrorClassName?: string;
  onKeyDown?: (event: KeyboardEvent<HTMLInputElement>) => void;
} & Omit<ComponentPropsWithRef<'div'>, 'onKeyDown'>) {
  const t = useTranslation();
  const input = useRef<HTMLInputElement>(null);
  const listId = useId();
  const mirror = useRef<HTMLDivElement>(null);
  const { entry } = useSyncExternalStore(view.subscribe, () => view.state);
  const text = entry?.text ?? contents;
  const here = entry?.place === place;
  const help = entry ? formulaHelp(entry.text, entry.caret) : null;
  // The list closed with Escape, until the entry changes.
  const [dismissed, setDismissed] = useState<Entry | null>(null);
  const [chosen, setChosen] = useState(0);
  const completion = here ? help?.completion : null;
  const offered =
    completion?.names ?? (here && entry.allFunctions ? functions : null);
  const names = entry !== dismissed ? offered : null;
  const highlighted = Math.min(chosen, (names?.length ?? 1) - 1);

  // The entry's text and caret, changed elsewhere, show here, and the
  // editor of the entry has the keyboard.
  useLayoutEffect(() => {
    const box = input.current!;
    if (!entry || !here) return;
    if (document.activeElement !== box) box.focus();
    if (caretOf(box) !== entry.caret) {
      box.setSelectionRange(entry.caret, entry.caret);
    }
    mirror.current?.scrollTo(box.scrollLeft, 0);
  }, [entry, here]);

  const choose = (name: string) => {
    view.complete(name);
    setChosen(0);
  };

  const keyDown = (event: KeyboardEvent<HTMLInputElement>) => {
    if (names) {
      const { key } = event;
      const move = key === 'ArrowDown' ? 1 : key === 'ArrowUp' ? -1 : 0;
      if (move !== 0) {
        setChosen((highlighted + move + names.length) % names.length);
      } else if (key === 'Tab' || key === 'Enter') {
        choose(names[highlighted]!);
      } else if (key === 'Escape') {
        setDismissed(entry);
      } else {
        onKeyDown?.(event);
        return;
      }
      // The grid around the cell's editor does not act on it too.
      event.preventDefault();
      event.stopPropagation();
      return;
    }
    onKeyDown?.(event);
  };

  const reportCaret = () => {
    if (input.current && here) view.moveCaret(caretOf(input.current));
  };

  return (
    <div
      {...around}
      className={classes(styles.around, className)}
      data-place={place}
    >
      <input
        ref={input}
        className={classes(styles.input, inputClassName)}
        data-formula={help && text.startsWith('=') ? '' : undefined}
        aria-label={label}
        aria-autocomplete="list"
        aria-expanded={names !== null}
        aria-controls={names ? listId : undefined}
        aria-activedescendant={names ? `${listId}-${highlighted}` : undefined}
        {...{ [EDITOR]: place }}
        value={text}
        onFocus={() => {
          if (view.state.entry?.place !== place) view.edit(place);
        }}
        onChange={(event) => {
          setChosen(0);
          view.type(event.target.value, caretOf(event.target));
        }}
        onSelect={reportCaret}
        onScroll={(event) =>
          mirror.current?.scrollTo(event.currentTarget.scrollLeft, 0)
        }
        onBlur={(event) => {
          if (!staysInEntry(event)) view.commit();
        }}
        onKeyDown={keyDown}
      />
      {help && text.startsWith('=') && (
        <div
          ref={mirror}
          className={classes(styles.mirror, mirrorClassName)}
          aria-hidden
        >
          {colored(text, help.references)}
        </div>
      )}
      {names && (
        <ul
          id={listId}
          className={styles.popup}
          role="listbox"
          aria-label={t('formula.functions')}
        >
          {names.map((name, index) => (
            <li
              key={name}
              id={`${listId}-${index}`}
              role="option"
              className={styles.option}
              aria-selected={index === highlighted}
              // The editor keeps the keyboard.
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => choose(name)}
            >
              {name}
            </li>
          ))}
        </ul>
      )}
      {here && !names && help?.hint && (
        <div
          className={styles.popup}
          role="note"
          aria-label={t('formula.arguments')}
        >
          <span className={styles.hint}>
            {`${help.hint.name}(`}
            {help.hint.arguments.map((argument, index) => (
              <span key={index}>
                {index > 0 && ', '}
                {index === help.hint!.current ? (
                  <strong>{argument}</strong>
                ) : (
                  argument
                )}
              </span>
            ))}
            {')'}
          </span>
        </div>
      )}
    </div>
  );
}

/** Where the caret is: the end of a selection that it moved. */
function caretOf(box: HTMLInputElement): number {
  return (
    (box.selectionDirection === 'backward'
      ? box.selectionStart
      : box.selectionEnd) ?? box.value.length
  );
}

/** The text of a formula with each reference in its color. */
function colored(
  text: string,
  references: ReturnType<typeof formulaHelp>['references'],
): ReactNode[] {
  const parts: ReactNode[] = [];
  let at = 0;
  for (const { start, end, color } of references) {
    parts.push(text.slice(at, start));
    parts.push(
      <span
        key={start}
        data-reference
        style={{ color: `var(--reference-${color})` }}
      >
        {text.slice(start, end)}
      </span>,
    );
    at = end;
  }
  parts.push(text.slice(at));
  return parts;
}

function classes(...names: (string | undefined)[]): string {
  return names.filter(Boolean).join(' ');
}
