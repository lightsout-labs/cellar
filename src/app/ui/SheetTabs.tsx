import * as ContextMenu from '@radix-ui/react-context-menu';
import {
  Check,
  Copy,
  Palette,
  Pencil,
  Plus,
  Trash2,
  type LucideIcon,
} from 'lucide-react';
import {
  useEffect,
  useReducer,
  useRef,
  useState,
  useSyncExternalStore,
  type CSSProperties,
  type MouseEvent,
  type PointerEvent,
} from 'react';
import { createPortal } from 'react-dom';
import type { SheetInfo, Workbook } from '../../engine/workbook';
import type { View } from '../grid/view';
import { Tooltip, useTranslation } from './controls';
import { dropIndex, TAB_COLORS } from './sheets';
import styles from './Shell.module.css';

const ICON_SIZE = 14;
/** How far, in pixels, the pointer moves before a press becomes a drag. */
const DRAG_START = 4;

/**
 * One tab for each sheet of the workbook, the one shown selected, and a
 * button that adds a sheet. A click shows a sheet, a double-click renames
 * it, a drag moves it, and a right-click opens the menu of its actions.
 */
export function SheetTabs({
  workbook,
  view,
}: {
  workbook: Workbook;
  view: View;
}) {
  const t = useTranslation();
  const { sheet: shown } = useSyncExternalStore(
    view.subscribe,
    () => view.state,
  );
  // The tabs change with the workbook.
  const [, changed] = useReducer((count: number) => count + 1, 0);
  useEffect(() => workbook.subscribe(changed), [workbook]);
  const sheets = workbook.sheets;
  const [renaming, setRenaming] = useState<string | null>(null);
  const strip = useRef<HTMLDivElement>(null);
  const press = useRef<{ index: number; x: number; dragging: boolean }>(null);
  const [drop, setDrop] = useState<{ from: number; to: number } | null>(null);

  const targetOf = (from: number, x: number) => {
    const tabs = [...strip.current!.querySelectorAll('[role="tab"]')].map(
      (tab) => tab.getBoundingClientRect(),
    );
    return dropIndex(tabs, from, x);
  };
  const onPointerDown = (index: number, event: PointerEvent) => {
    if (event.button !== 0) return;
    press.current = { index, x: event.clientX, dragging: false };
  };
  const onPointerMove = (event: PointerEvent<HTMLElement>) => {
    const current = press.current;
    if (!current) return;
    // The button was released somewhere the tab did not hear of it.
    if ((event.buttons & 1) === 0) {
      onPointerCancel();
      return;
    }
    if (!current.dragging) {
      if (Math.abs(event.clientX - current.x) < DRAG_START) return;
      current.dragging = true;
      event.currentTarget.setPointerCapture(event.pointerId);
    }
    setDrop({
      from: current.index,
      to: targetOf(current.index, event.clientX),
    });
  };
  const onPointerCancel = () => {
    press.current = null;
    setDrop(null);
  };
  const onPointerUp = (event: PointerEvent) => {
    const current = press.current;
    press.current = null;
    setDrop(null);
    if (!current?.dragging) return;
    const to = targetOf(current.index, event.clientX);
    const { id } = sheets[current.index]!;
    if (to !== current.index)
      view.execute({ type: 'setSheet', sheet: id, index: to });
  };

  return (
    <div className={styles.sheetTabs}>
      <div
        ref={strip}
        className={styles.tabStrip}
        role="tablist"
        aria-label={t('sheets')}
      >
        {sheets.map((sheet, index) =>
          sheet.id === renaming ? (
            <NameEditor
              key={sheet.id}
              workbook={workbook}
              sheet={sheet}
              onDone={(name) => {
                setRenaming(null);
                if (name !== null && name !== sheet.name) {
                  view.execute({ type: 'setSheet', sheet: sheet.id, name });
                }
              }}
            />
          ) : (
            <SheetTab
              key={sheet.id}
              sheet={sheet}
              selected={sheet.id === shown}
              only={sheets.length === 1}
              drop={
                drop && drop.to === index && drop.to !== drop.from
                  ? drop.to > drop.from
                    ? 'after'
                    : 'before'
                  : undefined
              }
              onPointerDown={(event) => onPointerDown(index, event)}
              onPointerMove={onPointerMove}
              onPointerUp={onPointerUp}
              onPointerCancel={onPointerCancel}
              onShow={() => view.showSheet(sheet.id)}
              onMouseDown={(event) => {
                // A formula being typed keeps the keyboard, to go on with
                // references on that sheet.
                if (view.state.entry) event.preventDefault();
              }}
              onRename={() => setRenaming(sheet.id)}
              onInsert={() =>
                view.execute({ type: 'addSheet', index: index + 1 })
              }
              onDuplicate={() =>
                view.execute({
                  type: 'addSheet',
                  index: index + 1,
                  copyOf: sheet.id,
                })
              }
              onDelete={() =>
                view.execute({ type: 'deleteSheet', sheet: sheet.id })
              }
              onColor={(color) =>
                view.execute({ type: 'setSheet', sheet: sheet.id, color })
              }
            />
          ),
        )}
      </div>
      <Tooltip text={t('sheet.new.tooltip')}>
        <button
          type="button"
          className={styles.iconButton}
          aria-label={t('sheet.new')}
          onClick={() => {
            const index = sheets.findIndex((sheet) => sheet.id === shown);
            view.execute({ type: 'addSheet', index: index + 1 });
          }}
        >
          <Plus size={ICON_SIZE} aria-hidden />
        </button>
      </Tooltip>
    </div>
  );
}

/** A sheet's tab, with the menu of the actions on its sheet. */
function SheetTab({
  sheet,
  selected,
  only,
  drop,
  onPointerDown,
  onPointerMove,
  onPointerUp,
  onPointerCancel,
  onShow,
  onMouseDown,
  onRename,
  onInsert,
  onDuplicate,
  onDelete,
  onColor,
}: {
  sheet: SheetInfo;
  selected: boolean;
  /** Whether it is the only sheet, which cannot be deleted. */
  only: boolean;
  /** Where the tab being dragged would go, beside this one. */
  drop: 'before' | 'after' | undefined;
  onPointerDown: (event: PointerEvent) => void;
  onPointerMove: (event: PointerEvent<HTMLElement>) => void;
  onPointerUp: (event: PointerEvent) => void;
  onPointerCancel: () => void;
  onShow: () => void;
  onMouseDown: (event: MouseEvent) => void;
  onRename: () => void;
  onInsert: () => void;
  onDuplicate: () => void;
  onDelete: () => void;
  onColor: (color: string | null) => void;
}) {
  const t = useTranslation();
  // The menu gives the keyboard back to the tab when it closes, which
  // would take it from the name editor that Rename opens.
  const renaming = useRef(false);
  return (
    <ContextMenu.Root>
      <Tooltip text={t('sheet.tooltip')}>
        <ContextMenu.Trigger asChild>
          <button
            type="button"
            role="tab"
            className={styles.sheetTab}
            aria-selected={selected}
            data-drop={drop}
            style={
              sheet.color
                ? ({ '--tab-color': sheet.color } as CSSProperties)
                : undefined
            }
            data-colored={sheet.color ? '' : undefined}
            onClick={onShow}
            onMouseDown={onMouseDown}
            onDoubleClick={onRename}
            onPointerDown={onPointerDown}
            onPointerMove={onPointerMove}
            onPointerUp={onPointerUp}
            onPointerCancel={onPointerCancel}
            onLostPointerCapture={onPointerCancel}
          >
            {sheet.name}
          </button>
        </ContextMenu.Trigger>
      </Tooltip>
      <ContextMenu.Portal>
        <ContextMenu.Content
          className={styles.menu}
          onCloseAutoFocus={(event) => {
            if (renaming.current) event.preventDefault();
            renaming.current = false;
          }}
        >
          <MenuItem icon={Plus} label={t('sheet.insert')} onSelect={onInsert} />
          <MenuItem
            icon={Copy}
            label={t('sheet.duplicate')}
            onSelect={onDuplicate}
          />
          <MenuItem
            icon={Pencil}
            label={t('sheet.rename')}
            onSelect={() => {
              renaming.current = true;
              onRename();
            }}
          />
          <MenuItem
            icon={Trash2}
            label={t('sheet.delete')}
            disabled={only}
            onSelect={onDelete}
          />
          <ContextMenu.Separator className={styles.menuSeparator} />
          <ContextMenu.Sub>
            <ContextMenu.SubTrigger className={styles.menuItem}>
              <Palette size={ICON_SIZE} aria-hidden />
              {t('sheet.color')}
            </ContextMenu.SubTrigger>
            <ContextMenu.Portal>
              <ContextMenu.SubContent className={styles.menu} sideOffset={4}>
                {TAB_COLORS.map(([name, color]) => (
                  <ContextMenu.Item
                    key={color}
                    className={styles.menuItem}
                    onSelect={() => onColor(color)}
                  >
                    <span
                      className={styles.swatch}
                      style={{ background: color }}
                      aria-hidden
                    />
                    {t(name)}
                    {sheet.color === color && (
                      <Check
                        className={styles.menuCheck}
                        size={ICON_SIZE}
                        aria-hidden
                      />
                    )}
                  </ContextMenu.Item>
                ))}
                <ContextMenu.Separator className={styles.menuSeparator} />
                <ContextMenu.Item
                  className={styles.menuItem}
                  disabled={!sheet.color}
                  onSelect={() => onColor(null)}
                >
                  <span className={styles.swatch} aria-hidden />
                  {t('color.none')}
                </ContextMenu.Item>
              </ContextMenu.SubContent>
            </ContextMenu.Portal>
          </ContextMenu.Sub>
        </ContextMenu.Content>
      </ContextMenu.Portal>
    </ContextMenu.Root>
  );
}

function MenuItem({
  icon: Icon,
  label,
  disabled,
  onSelect,
}: {
  icon: LucideIcon;
  label: string;
  disabled?: boolean;
  onSelect: () => void;
}) {
  return (
    <ContextMenu.Item
      className={styles.menuItem}
      disabled={disabled}
      onSelect={onSelect}
    >
      <Icon size={ICON_SIZE} aria-hidden />
      {label}
    </ContextMenu.Item>
  );
}

/**
 * The box in place of a tab where its sheet is renamed. Enter keeps a name
 * Excel accepts, and says what is wrong with any other; leaving the box
 * keeps a name Excel accepts, as Excel does; Escape keeps the name the sheet
 * had.
 */
function NameEditor({
  workbook,
  sheet,
  onDone,
}: {
  workbook: Workbook;
  sheet: SheetInfo;
  onDone: (name: string | null) => void;
}) {
  const t = useTranslation();
  const [name, setName] = useState(sheet.name);
  // The box may lose the keyboard as it closes, which must not end it twice.
  const ended = useRef(false);
  const end = (name: string | null) => {
    if (ended.current) return;
    ended.current = true;
    onDone(name);
  };
  // Shown above the box, outside the tab strip, which would clip it.
  const [problem, setProblem] = useState<{
    text: string;
    left: number;
    bottom: number;
  } | null>(null);
  return (
    <span className={styles.nameEditor}>
      <input
        className={styles.nameInput}
        aria-label={t('sheet.name')}
        aria-invalid={problem !== null}
        value={name}
        size={Math.max(name.length, 4)}
        autoFocus
        onFocus={(event) => event.target.select()}
        onChange={(event) => {
          setName(event.target.value);
          setProblem(null);
        }}
        onBlur={() =>
          end(workbook.sheetNameProblem(name, sheet.id) ? null : name)
        }
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault();
            end(null);
          } else if (event.key === 'Enter') {
            event.preventDefault();
            const found = workbook.sheetNameProblem(name, sheet.id);
            if (!found) {
              end(name);
              return;
            }
            const box = event.currentTarget.getBoundingClientRect();
            setProblem({
              text: t(`sheetName.${found}`),
              left: box.left,
              bottom: window.innerHeight - box.top + 4,
            });
          }
        }}
      />
      {problem &&
        createPortal(
          <span
            className={styles.nameProblem}
            role="alert"
            style={{ left: problem.left, bottom: problem.bottom }}
          >
            {problem.text}
          </span>,
          document.body,
        )}
    </span>
  );
}
