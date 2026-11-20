# Fonts in this folder (the "Times New Roman" profile: newtxtext + newtxmath)

The `vtx-n-*.otf` fonts are OpenType (CFF) conversions of the fonts LaTeX uses
with `\usepackage{newtxtext,newtxmath}`: each converted font is the TFM font
of the same name, its characters composed exactly as the virtual font (VF)
draws them from these Type 1 fonts (tools/vffonts.py in the VisualTeX layout
sources):

- TeX Gyre Termes and Heros (ztmr, ztmb, ztmri, ztmbi as modified in newtx;
  qhvr, qhvb, qhvri, qhvbi). Copyright 2006-2009 B. Jackowski and J. M. Nowacki
  (on behalf of TeX users groups); newtx modifications (c) 2015-2023 Michael
  Sharpe. GUST Font License (http://tug.org/fonts/licenses/GUST-FONT-LICENSE.txt),
  an LPPL-type license; modified versions are renamed.
- txfonts and the newtx math fonts derived from them (txsys, txsym, txsyc, txexs,
  txexas, txex-bar, txmiaX, txmiaSTbb, txbsys, txbexs, txbex-bar, txbmiaX, txtt,
  txttsc, txbtt, txbttsc; NewTXMI, NewTXBMI and their 5/7 sizes, also derived
  from GNU FreeFont). Copyright Young Ryu; modifications (c) Michael Sharpe.
  GNU General Public License. The Type 1 sources and the converter are
  available from CTAN (fonts/newtx, fonts/txfonts) and the VisualTeX sources.
- STIX script letters (stxscr). Copyright (c) 2001-2011 the STI Pub Companies.
  SIL Open Font License 1.1.

The converted fonts are renamed ("VTX N <tfm name>"). Each TeX character code c
is mapped to U+E000+c (Private Use Area).
