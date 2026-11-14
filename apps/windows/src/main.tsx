import { bootVisualTeX } from "../../shared/visualtex-app/src/runtime/boot";

bootVisualTeX(() => import("./desktop/main"));
