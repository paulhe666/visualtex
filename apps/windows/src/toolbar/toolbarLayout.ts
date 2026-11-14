// The formula toolbar's groups (the tabs of the tools view) as the user arranged them: their order, their names and
// what each holds. A built-in group follows the app's own content until the user edits it (then its items are its
// own); a group the user made always has its own. A removed built-in group can be restored. Shared by every toolbar
// on the page (the sidebar, the classic bottom toolbar, the tiles panel: a drag goes from one to another) and kept in
// step across the app's windows.
import { create } from "zustand";
import {
  calculusCommandIds,
  commandRegistry,
  commonCommandIds,
} from "../autocomplete/commandRegistry";
import { writeLocalStorage } from "../../../shared/visualtex-app/src/runtime/safeStorage";

export const BUILT_IN_GROUP_IDS = [
  "common",
  "structure",
  "calculus",
  "matrix",
  "relation",
  "greek",
  "arrow",
  "physics",
  "set",
] as const;
export type BuiltInGroupId = (typeof BUILT_IN_GROUP_IDS)[number];

/** One thing a group holds: a command of the registry (custom symbols too), a formula of the user's, the matrix builder. */
export type ToolbarItem =
  | { kind: "command"; id: string }
  | { kind: "formula"; id: string; latex: string }
  | { kind: "matrix" };

export interface ToolbarGroup {
  id: string;
  /** the built-in group it is (its default content and name) */
  builtIn?: BuiltInGroupId;
  /** the user's name (none: the built-in group's own, in the interface's language) */
  name?: string;
  /** the user's items (none: the built-in group's default content, which follows the app) */
  items?: ToolbarItem[];
}

export interface ToolbarLayout {
  version: 1;
  groups: ToolbarGroup[];
}

export const TOOLBAR_LAYOUT_STORAGE_KEY = "visualtex-toolbar-layout-v1";
const LEGACY_COMMON_KEY = "visualtex-common-toolbar-command-ids-v2";
const LEGACY_COMMON_KEY_V1 = "visualtex-common-toolbar-command-ids-v1";
export const COMMON_GROUP_COMMAND_LIMIT = 45;
const MAX_GROUPS = 40;
const MAX_GROUP_ITEMS = 300;
const MAX_GROUP_NAME = 24;

// Package shorthands remain available through autocomplete and raw LaTeX, but
// the toolbar should expose each semantic template only once. Prefer the
// primitive templates here: they render more compactly and give each visible
// placeholder a direct, predictable caret stop.
export const hiddenToolbarCommandIds = new Set([
  "time-ordering",
  "shortcut-bra",
  "shortcut-ket",
  "shortcut-expval",
  "shortcut-comm",
  "shortcut-acomm",
  "shortcut-braket",
  "shortcut-ketbra",
  "shortcut-mel",
]);

const registryById = new Map(commandRegistry.map((command) => [command.id, command]));

/** The Common group's commands: the given ids first (known, once), then the defaults, up to the limit. */
export function normalizeCommonToolbarCommandIds(value: unknown) {
  const result: string[] = [];
  const append = (id: unknown) => {
    if (typeof id !== "string" || !registryById.has(id) || result.includes(id) || result.length >= COMMON_GROUP_COMMAND_LIMIT) return;
    result.push(id);
  };
  if (Array.isArray(value)) value.forEach(append);
  commonCommandIds.forEach(append);
  commandRegistry.filter((command) => !hiddenToolbarCommandIds.has(command.id)).forEach((command) => append(command.id));
  return result.slice(0, COMMON_GROUP_COMMAND_LIMIT);
}

function migrateLegacyCommonIds(value: unknown) {
  if (!Array.isArray(value)) return value;
  return value.map((id) => (id === "notin" ? "times" : id === "leftarrow" ? "div" : id));
}

/** A built-in group's content as the app ships it. */
export function defaultGroupItems(id: BuiltInGroupId): ToolbarItem[] {
  const ids =
    id === "common"
      ? normalizeCommonToolbarCommandIds(null)
      : id === "calculus"
        ? calculusCommandIds.filter((commandId) => registryById.has(commandId))
        : commandRegistry.filter((command) => command.category === id && !hiddenToolbarCommandIds.has(command.id)).map((command) => command.id);
  const seen = new Set<string>();
  const commands: ToolbarItem[] = ids.filter((commandId) => !seen.has(commandId) && Boolean(seen.add(commandId))).map((commandId) => ({ kind: "command", id: commandId }));
  return id === "matrix" ? [{ kind: "matrix" }, ...commands] : commands;
}

export function defaultToolbarLayout(): ToolbarLayout {
  return { version: 1, groups: BUILT_IN_GROUP_IDS.map((id) => ({ id, builtIn: id })) };
}

/** What a group holds now. */
export function groupItems(group: ToolbarGroup): ToolbarItem[] {
  return group.items ?? (group.builtIn ? defaultGroupItems(group.builtIn) : []);
}

/** Two items are the same thing (a group holds a thing once). */
export function sameToolbarItem(a: ToolbarItem, b: ToolbarItem) {
  if (a.kind !== b.kind) return false;
  if (a.kind === "command" && b.kind === "command") return a.id === b.id;
  if (a.kind === "formula" && b.kind === "formula") return a.latex.trim() === b.latex.trim();
  return a.kind === "matrix";
}

function newId(prefix: string) {
  return `${prefix}-${globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`}`;
}

function normalizeItem(value: unknown): ToolbarItem | null {
  if (!value || typeof value !== "object") return null;
  const item = value as Record<string, unknown>;
  if (item.kind === "matrix") return { kind: "matrix" };
  if (item.kind === "command" && typeof item.id === "string" && item.id) return { kind: "command", id: item.id };
  if (item.kind === "formula" && typeof item.latex === "string" && item.latex.trim()) {
    return { kind: "formula", id: typeof item.id === "string" && item.id ? item.id : newId("formula"), latex: item.latex.trim() };
  }
  return null;
}

function normalizeItems(value: unknown): ToolbarItem[] {
  const items: ToolbarItem[] = [];
  if (!Array.isArray(value)) return items;
  for (const raw of value) {
    const item = normalizeItem(raw);
    if (item && !items.some((existing) => sameToolbarItem(existing, item))) items.push(item);
    if (items.length >= MAX_GROUP_ITEMS) break;
  }
  return items;
}

export function normalizeToolbarLayout(value: unknown): ToolbarLayout | null {
  if (!value || typeof value !== "object" || !Array.isArray((value as ToolbarLayout).groups)) return null;
  const groups: ToolbarGroup[] = [];
  const ids = new Set<string>();
  for (const raw of (value as ToolbarLayout).groups) {
    if (!raw || typeof raw !== "object") continue;
    const builtIn = (BUILT_IN_GROUP_IDS as readonly string[]).includes(raw.builtIn as string) ? (raw.builtIn as BuiltInGroupId) : undefined;
    const id = builtIn ?? (typeof raw.id === "string" && raw.id ? raw.id : newId("group"));
    if (ids.has(id)) continue;
    ids.add(id);
    const name = typeof raw.name === "string" && raw.name.trim() ? raw.name.trim().slice(0, MAX_GROUP_NAME) : undefined;
    // a group of the user's always has its items and a name
    if (!builtIn && !name) continue;
    const group: ToolbarGroup = { id, ...(builtIn ? { builtIn } : {}), ...(name ? { name } : {}) };
    if (raw.items !== undefined || !builtIn) group.items = normalizeItems(raw.items);
    groups.push(group);
    if (groups.length >= MAX_GROUPS) break;
  }
  return { version: 1, groups };
}

function readStorage(key: string) {
  try {
    return globalThis.localStorage?.getItem(key) ?? null;
  } catch {
    return null;
  }
}

/** The saved layout; the first time, the shipped one with the user's Common commands of before (if they edited it). */
export function loadToolbarLayout(): ToolbarLayout {
  const stored = readStorage(TOOLBAR_LAYOUT_STORAGE_KEY);
  if (stored) {
    try {
      const layout = normalizeToolbarLayout(JSON.parse(stored));
      if (layout) return layout;
    } catch {
      // the shipped layout below
    }
  }
  const layout = defaultToolbarLayout();
  const legacy = readStorage(LEGACY_COMMON_KEY) ?? readStorage(LEGACY_COMMON_KEY_V1);
  if (legacy) {
    try {
      const ids = normalizeCommonToolbarCommandIds(migrateLegacyCommonIds(JSON.parse(legacy)));
      const defaults = normalizeCommonToolbarCommandIds(null);
      if (ids.join("\n") !== defaults.join("\n")) {
        layout.groups[0] = { ...layout.groups[0], items: ids.map((id) => ({ kind: "command", id })) };
      }
    } catch {
      // keep the shipped Common group
    }
  }
  return layout;
}

function persist(layout: ToolbarLayout) {
  writeLocalStorage(TOOLBAR_LAYOUT_STORAGE_KEY, JSON.stringify(layout));
}


interface ToolbarLayoutState {
  layout: ToolbarLayout;
  /** the toolbars are being arranged (clicks do not insert; groups and items can be moved, renamed, removed) */
  editing: boolean;
  setEditing: (editing: boolean) => void;
  addGroup: (name: string) => string;
  renameGroup: (groupId: string, name: string) => void;
  removeGroup: (groupId: string) => void;
  moveGroup: (groupId: string, toIndex: number) => void;
  restoreGroup: (builtIn: BuiltInGroupId) => void;
  resetGroupItems: (groupId: string) => void;
  resetAll: () => void;
  clearAll: () => void;
  /** put an item into a group at an index (a thing the group holds already moves there) */
  insertItem: (groupId: string, index: number, item: ToolbarItem) => void;
  moveItem: (fromGroupId: string, fromIndex: number, toGroupId: string, toIndex: number) => void;
  removeItem: (groupId: string, index: number) => void;
}

function withGroups(_state: { layout: ToolbarLayout }, groups: ToolbarGroup[]) {
  const layout: ToolbarLayout = { version: 1, groups };
  persist(layout);
  return { layout };
}

/** A copy of the groups with one group's items replaced (a built-in group now holds its own). */
function setItems(groups: ToolbarGroup[], groupId: string, items: ToolbarItem[]) {
  return groups.map((group) => (group.id === groupId ? { ...group, items: items.slice(0, MAX_GROUP_ITEMS) } : group));
}

export const useToolbarLayoutStore = create<ToolbarLayoutState>()((set, get) => ({
  layout: loadToolbarLayout(),
  editing: false,
  setEditing: (editing) => set({ editing }),
  addGroup: (name) => {
    const id = newId("group");
    const trimmed = name.trim().slice(0, MAX_GROUP_NAME) || "新分组";
    set((state) => (state.layout.groups.length >= MAX_GROUPS ? state : withGroups(state, [...state.layout.groups, { id, name: trimmed, items: [] }])));
    return id;
  },
  renameGroup: (groupId, name) =>
    set((state) =>
      withGroups(
        state,
        state.layout.groups.map((group) => {
          if (group.id !== groupId) return group;
          const trimmed = name.trim().slice(0, MAX_GROUP_NAME);
          // an emptied name: a built-in group takes its own again; a user group keeps its old one
          if (!trimmed) return group.builtIn ? { ...group, name: undefined } : group;
          return { ...group, name: trimmed };
        }),
      ),
    ),
  removeGroup: (groupId) => set((state) => withGroups(state, state.layout.groups.filter((group) => group.id !== groupId))),
  moveGroup: (groupId, toIndex) =>
    set((state) => {
      const groups = [...state.layout.groups];
      const from = groups.findIndex((group) => group.id === groupId);
      if (from < 0) return state;
      const [group] = groups.splice(from, 1);
      groups.splice(Math.max(0, Math.min(groups.length, toIndex > from ? toIndex - 1 : toIndex)), 0, group);
      return withGroups(state, groups);
    }),
  restoreGroup: (builtIn) =>
    set((state) => {
      if (state.layout.groups.some((group) => group.id === builtIn)) return state;
      // back where the app puts it among the built-in groups still there
      const order = BUILT_IN_GROUP_IDS.indexOf(builtIn);
      const groups = [...state.layout.groups];
      const after = groups.findIndex((group) => group.builtIn !== undefined && BUILT_IN_GROUP_IDS.indexOf(group.builtIn) > order);
      groups.splice(after < 0 ? groups.length : after, 0, { id: builtIn, builtIn });
      return withGroups(state, groups);
    }),
  resetGroupItems: (groupId) =>
    set((state) => withGroups(state, state.layout.groups.map((group) => (group.id === groupId && group.builtIn ? { id: group.id, builtIn: group.builtIn, ...(group.name ? { name: group.name } : {}) } : group)))),
  resetAll: () => set((state) => withGroups(state, defaultToolbarLayout().groups)),
  clearAll: () => set((state) => withGroups(state, [])),
  insertItem: (groupId, index, item) =>
    set((state) => {
      const group = state.layout.groups.find((candidate) => candidate.id === groupId);
      if (!group) return state;
      const items = groupItems(group);
      const existing = items.findIndex((candidate) => sameToolbarItem(candidate, item));
      const next = items.filter((_, i) => i !== existing);
      const at = Math.max(0, Math.min(next.length, existing >= 0 && existing < index ? index - 1 : index));
      next.splice(at, 0, existing >= 0 ? items[existing] : item);
      return withGroups(state, setItems(state.layout.groups, groupId, next));
    }),
  moveItem: (fromGroupId, fromIndex, toGroupId, toIndex) => {
    const groups = get().layout.groups;
    const from = groups.find((group) => group.id === fromGroupId);
    if (!from) return;
    const fromItems = groupItems(from);
    const item = fromItems[fromIndex];
    if (!item) return;
    if (fromGroupId === toGroupId) {
      get().insertItem(toGroupId, toIndex, item);
      return;
    }
    set((state) => {
      const target = state.layout.groups.find((group) => group.id === toGroupId);
      if (!target) return state;
      let next = setItems(state.layout.groups, fromGroupId, fromItems.filter((_, i) => i !== fromIndex));
      const targetItems = groupItems(target).filter((candidate) => !sameToolbarItem(candidate, item));
      targetItems.splice(Math.max(0, Math.min(targetItems.length, toIndex)), 0, item);
      next = setItems(next, toGroupId, targetItems);
      return withGroups(state, next);
    });
  },
  removeItem: (groupId, index) =>
    set((state) => {
      const group = state.layout.groups.find((candidate) => candidate.id === groupId);
      if (!group) return state;
      return withGroups(state, setItems(state.layout.groups, groupId, groupItems(group).filter((_, i) => i !== index)));
    }),
}));

// another window changed the layout: take it
if (typeof window !== "undefined") {
  window.addEventListener("storage", (event) => {
    if (event.key !== TOOLBAR_LAYOUT_STORAGE_KEY) return;
    useToolbarLayoutStore.setState({ layout: loadToolbarLayout() });
  });
}
