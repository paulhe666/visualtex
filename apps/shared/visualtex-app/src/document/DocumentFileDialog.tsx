// A question or a report about a document's file (document mode): a title, the lines it is about and its buttons.
// It stays until one of its buttons is pressed.
import { createPortal } from "react-dom";

export interface DocumentFileDialogAction {
  label: string;
  run: () => void;
  primary?: boolean;
}

export function DocumentFileDialog(props: { title: string; lines: string[]; actions: DocumentFileDialogAction[] }) {
  return createPortal(
    <div className="modal-backdrop vtx-file-dialog-backdrop" role="presentation">
      <section className="dialog vtx-file-dialog" role="alertdialog" aria-modal="true" aria-label={props.title}>
        <h2 className="vtx-file-dialog-title">{props.title}</h2>
        {props.lines.length > 0 && (
          <ul className="vtx-file-dialog-lines">
            {props.lines.map((line, index) => (
              <li key={index}>{line}</li>
            ))}
          </ul>
        )}
        <div className="vtx-file-dialog-actions">
          {props.actions.map((action) => (
            <button key={action.label} type="button" className={action.primary ? "primary-small-button" : "secondary-button"} autoFocus={action.primary} onClick={action.run}>
              {action.label}
            </button>
          ))}
        </div>
      </section>
    </div>,
    document.body,
  );
}
