VisualTeX private MathPage runtime staging directory.

Place a legally obtained x64 MathPage.wll in this directory as:
  MathPage.wll

The Windows bundle maps this directory to:
  <VisualTeX install root>\mathtype-runtime\

The WLL is loaded only by visualtex-windows-office-bridge.exe for isolated
MTEF -> native MathType WMF preview generation. It is not copied into Word's
STARTUP directory and is not registered as a Word add-in.

MathPage.wll is intentionally ignored by Git and is not included in this
source tree. Redistribution requires the appropriate rights from its owner.
