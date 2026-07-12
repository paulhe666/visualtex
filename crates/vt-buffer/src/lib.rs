use std::collections::{HashSet, VecDeque};
use std::path::PathBuf;

use ropey::Rope;
use vt_protocol::{AppliedEdit, EditOrigin, FileId, OperationId, Revision, TextEdit};

#[derive(Debug, thiserror::Error, PartialEq, Eq)]
pub enum BufferError {
    #[error("stale revision: expected {expected:?}, received {received:?}")]
    StaleRevision {
        expected: Revision,
        received: Revision,
    },
    #[error("invalid byte range {start}..{end} for document with {len} bytes")]
    InvalidRange {
        start: usize,
        end: usize,
        len: usize,
    },
    #[error("byte offset {0} is not a UTF-8 character boundary")]
    InvalidUtf8Boundary(usize),
    #[error("operation {0:?} was already applied")]
    DuplicateOperation(OperationId),
    #[error("undo history is empty")]
    NothingToUndo,
    #[error("redo history is empty")]
    NothingToRedo,
}

#[derive(Clone, Debug)]
struct HistoryEntry {
    forward: TextEdit,
    inverse_start: usize,
    inverse_end: usize,
    inverse_text: String,
}

#[derive(Clone, Debug)]
pub struct DocumentBuffer {
    pub file_id: FileId,
    pub path: PathBuf,
    pub revision: Revision,
    text: Rope,
    pub dirty: bool,
    applied_operations: HashSet<OperationId>,
    undo: VecDeque<HistoryEntry>,
    redo: VecDeque<HistoryEntry>,
    history_limit: usize,
}

impl DocumentBuffer {
    pub fn new(path: PathBuf, text: impl AsRef<str>) -> Self {
        Self::with_file_id(FileId::new(), path, text)
    }

    pub fn with_file_id(file_id: FileId, path: PathBuf, text: impl AsRef<str>) -> Self {
        Self {
            file_id,
            path,
            revision: Revision::default(),
            text: Rope::from_str(text.as_ref()),
            dirty: false,
            applied_operations: HashSet::new(),
            undo: VecDeque::new(),
            redo: VecDeque::new(),
            history_limit: 10_000,
        }
    }

    pub fn text(&self) -> String {
        self.text.to_string()
    }

    pub fn len_bytes(&self) -> usize {
        self.text.len_bytes()
    }

    pub fn set_history_limit(&mut self, limit: usize) {
        self.history_limit = limit.max(1);
        self.trim_history();
    }

    pub fn mark_saved(&mut self) {
        self.dirty = false;
    }

    pub fn mark_dirty(&mut self) {
        self.dirty = true;
    }

    pub fn apply(&mut self, edit: TextEdit) -> Result<AppliedEdit, BufferError> {
        self.apply_internal(edit, true, true)
    }

    pub fn apply_external_text(&mut self, replacement: String) -> Result<AppliedEdit, BufferError> {
        let edit = TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::ExternalFileChange,
            file_id: self.file_id,
            base_revision: self.revision,
            start_byte: 0,
            end_byte: self.len_bytes(),
            replacement,
        };
        self.apply(edit)
    }

    pub fn undo(&mut self) -> Result<AppliedEdit, BufferError> {
        let entry = self.undo.pop_back().ok_or(BufferError::NothingToUndo)?;
        let current = self.slice_bytes(entry.inverse_start, entry.inverse_end)?;
        let edit = TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::UndoRedo,
            file_id: self.file_id,
            base_revision: self.revision,
            start_byte: entry.inverse_start,
            end_byte: entry.inverse_end,
            replacement: entry.inverse_text.clone(),
        };
        let applied = self.apply_internal(edit, false, false)?;
        self.redo.push_back(HistoryEntry {
            forward: entry.forward,
            inverse_start: entry.inverse_start,
            inverse_end: entry.inverse_start + entry.inverse_text.len(),
            inverse_text: current,
        });
        Ok(applied)
    }

    pub fn redo(&mut self) -> Result<AppliedEdit, BufferError> {
        let entry = self.redo.pop_back().ok_or(BufferError::NothingToRedo)?;
        let current = self.slice_bytes(entry.inverse_start, entry.inverse_end)?;
        let edit = TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::UndoRedo,
            file_id: self.file_id,
            base_revision: self.revision,
            start_byte: entry.inverse_start,
            end_byte: entry.inverse_end,
            replacement: entry.inverse_text.clone(),
        };
        let applied = self.apply_internal(edit, false, false)?;
        self.undo.push_back(HistoryEntry {
            forward: entry.forward,
            inverse_start: entry.inverse_start,
            inverse_end: entry.inverse_start + entry.inverse_text.len(),
            inverse_text: current,
        });
        Ok(applied)
    }

    fn apply_internal(
        &mut self,
        edit: TextEdit,
        record_history: bool,
        clear_redo: bool,
    ) -> Result<AppliedEdit, BufferError> {
        if edit.file_id != self.file_id {
            return Err(BufferError::InvalidRange {
                start: edit.start_byte,
                end: edit.end_byte,
                len: self.len_bytes(),
            });
        }
        if edit.base_revision != self.revision {
            return Err(BufferError::StaleRevision {
                expected: self.revision,
                received: edit.base_revision,
            });
        }
        if self.applied_operations.contains(&edit.operation_id) {
            return Err(BufferError::DuplicateOperation(edit.operation_id));
        }
        self.validate_range(edit.start_byte, edit.end_byte)?;

        let replaced_text = self.slice_bytes(edit.start_byte, edit.end_byte)?;
        let start_char = self.text.byte_to_char(edit.start_byte);
        let end_char = self.text.byte_to_char(edit.end_byte);
        self.text.remove(start_char..end_char);
        self.text.insert(start_char, &edit.replacement);

        let old_revision = self.revision;
        self.revision = self.revision.next();
        self.dirty = true;
        self.applied_operations.insert(edit.operation_id);

        if record_history {
            self.undo.push_back(HistoryEntry {
                forward: edit.clone(),
                inverse_start: edit.start_byte,
                inverse_end: edit.start_byte + edit.replacement.len(),
                inverse_text: replaced_text,
            });
            self.trim_history();
            if clear_redo {
                self.redo.clear();
            }
        }

        Ok(AppliedEdit {
            operation_id: edit.operation_id,
            file_id: edit.file_id,
            old_revision,
            new_revision: self.revision,
            start_byte: edit.start_byte,
            old_end_byte: edit.end_byte,
            new_end_byte: edit.start_byte + edit.replacement.len(),
        })
    }

    fn trim_history(&mut self) {
        while self.undo.len() > self.history_limit {
            self.undo.pop_front();
        }
        while self.redo.len() > self.history_limit {
            self.redo.pop_front();
        }
    }

    fn validate_range(&self, start: usize, end: usize) -> Result<(), BufferError> {
        let len = self.len_bytes();
        if start > end || end > len {
            return Err(BufferError::InvalidRange { start, end, len });
        }
        if !self.is_char_boundary(start) {
            return Err(BufferError::InvalidUtf8Boundary(start));
        }
        if !self.is_char_boundary(end) {
            return Err(BufferError::InvalidUtf8Boundary(end));
        }
        Ok(())
    }

    fn is_char_boundary(&self, byte: usize) -> bool {
        if byte == self.len_bytes() {
            return true;
        }
        let char_index = self.text.byte_to_char(byte);
        self.text.char_to_byte(char_index) == byte
    }

    fn slice_bytes(&self, start: usize, end: usize) -> Result<String, BufferError> {
        self.validate_range(start, end)?;
        Ok(self
            .text
            .slice(self.text.byte_to_char(start)..self.text.byte_to_char(end))
            .to_string())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use proptest::prelude::*;

    fn edit(buffer: &DocumentBuffer, start: usize, end: usize, replacement: &str) -> TextEdit {
        TextEdit {
            operation_id: OperationId::new(),
            origin: EditOrigin::SourceEditor,
            file_id: buffer.file_id,
            base_revision: buffer.revision,
            start_byte: start,
            end_byte: end,
            replacement: replacement.to_owned(),
        }
    }

    #[test]
    fn applies_unicode_edit_on_byte_boundaries() {
        let mut buffer = DocumentBuffer::new("main.tex".into(), "你好, world");
        buffer
            .apply(edit(&buffer, 0, "你好".len(), "您好"))
            .unwrap();
        assert_eq!(buffer.text(), "您好, world");
        assert_eq!(buffer.revision, Revision(1));
    }

    #[test]
    fn rejects_middle_of_unicode_scalar() {
        let mut buffer = DocumentBuffer::new("main.tex".into(), "你");
        let error = buffer.apply(edit(&buffer, 1, 1, "x")).unwrap_err();
        assert_eq!(error, BufferError::InvalidUtf8Boundary(1));
    }

    #[test]
    fn rejects_stale_revision() {
        let mut buffer = DocumentBuffer::new("main.tex".into(), "abc");
        let stale = edit(&buffer, 0, 1, "A");
        buffer.apply(edit(&buffer, 1, 2, "B")).unwrap();
        assert!(matches!(
            buffer.apply(stale),
            Err(BufferError::StaleRevision { .. })
        ));
    }

    #[test]
    fn undo_and_redo_round_trip() {
        let mut buffer = DocumentBuffer::new("main.tex".into(), "abc");
        buffer.apply(edit(&buffer, 1, 2, "XYZ")).unwrap();
        assert_eq!(buffer.text(), "aXYZc");
        buffer.undo().unwrap();
        assert_eq!(buffer.text(), "abc");
        buffer.redo().unwrap();
        assert_eq!(buffer.text(), "aXYZc");
    }

    proptest! {
        #[test]
        fn arbitrary_unicode_edit_matches_std_string_and_history(
            original_chars in proptest::collection::vec(any::<char>(), 0..80),
            replacement_chars in proptest::collection::vec(any::<char>(), 0..30),
            start_seed in any::<usize>(),
            end_seed in any::<usize>(),
        ) {
            let original = original_chars.iter().collect::<String>();
            let replacement = replacement_chars.iter().collect::<String>();
            let start_char = start_seed % (original_chars.len() + 1);
            let remaining = original_chars.len() - start_char;
            let end_char = start_char + end_seed % (remaining + 1);
            let byte_offsets = original
                .char_indices()
                .map(|(offset, _)| offset)
                .chain(std::iter::once(original.len()))
                .collect::<Vec<_>>();
            let start_byte = byte_offsets[start_char];
            let end_byte = byte_offsets[end_char];

            let mut expected = original.clone();
            expected.replace_range(start_byte..end_byte, &replacement);

            let mut buffer = DocumentBuffer::new("property.tex".into(), &original);
            buffer
                .apply(edit(&buffer, start_byte, end_byte, &replacement))
                .unwrap();
            prop_assert_eq!(buffer.text(), expected.clone());
            buffer.undo().unwrap();
            prop_assert_eq!(buffer.text(), original);
            buffer.redo().unwrap();
            prop_assert_eq!(buffer.text(), expected);
        }
    }
}
