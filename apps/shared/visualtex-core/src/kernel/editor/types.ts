import { Selector } from '../public/commands';
export type SelectorPrivate = Selector;

export type CommandRegistry<T> = Partial<{
  [K in SelectorPrivate]: { fn: (...args: unknown[]) => boolean } & T;
}>;
