// the stylesheets in the page itself: index.html links them before any script runs (the Windows build checks it)
import "../../shared/visualtex-app/src/styles/index.ts";
import { bootVisualTeX } from "../../shared/visualtex-app/src/runtime/boot";

bootVisualTeX(() => import("./desktop/main"));
