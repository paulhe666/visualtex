# Active P07 implementation continuation

2026-09-13 14:51 +08:00. Same checkout/workspace; HEAD stays 00b48e17b736c7e57a0a72e19e271df82906f89d. No engine Git commit/push.

CP06-native-ime-verified and CP07-radicals are complete immutable non-Git checkpoints with audits. Current-source builds and real Windows IME were verified, followed by the rooted structural slice and 213 native tests + Native/WASM + sandbox/current-host checks. Source stayed unchanged during final verification.

This implementation lane now owns P07 delimiter changes in core model/parser/transactions/positions/projection/source_edit/fonts/stretch/layout and existing engine/web/host/editor-host.ts plus sandbox. Do not introduce a second host or change production entries. Independent reviews should use separately named tests and not overwrite these files during construction. Existing user Word and installed apps remain protected. Next archive must not overwrite CP07-radicals.
