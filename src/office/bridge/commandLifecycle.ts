export interface OfficeCommandEvent {
  completed?: () => void;
}

export function startOfficeDialogCommand(
  run: () => Promise<void>,
  event?: OfficeCommandEvent,
) {
  const operation = run();
  event?.completed?.();
  void operation;
}
