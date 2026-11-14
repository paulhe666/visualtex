// The "Times New Roman" profile (newtxtext + newtxmath) of the layout engine:
// its data and fonts, loaded only when that letter font is chosen.
import tfm from "../../../tex-layout/data/profiles/ntx/tfm.json";
import meanings from "../../../tex-layout/data/profiles/ntx/latex-meanings.json";
import templates from "../../../tex-layout/data/profiles/ntx/latex-templates.json";
import onearg from "../../../tex-layout/data/profiles/ntx/latex-onearg.json";
import textFonts from "../../../tex-layout/data/profiles/ntx/text-fonts.json";
import fonts from "../../../tex-layout/data/profiles/ntx/fonts.json";
import profile from "../../../tex-layout/data/profiles/ntx/profile.json";
import "../../../tex-layout/data/profiles/ntx/fonts/fonts.css";

export default { tfm, meanings, templates, onearg, textFonts, fonts, profile };
